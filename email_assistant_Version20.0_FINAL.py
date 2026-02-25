"""
Version 20.0 FINAL – Email Assistant - INTELLIGENT PRODUCTION VERSION

NEW FEATURES:
✅ Smart Email-to-Checklist Integration (keyword parsing)
✅ Anchored Date Logic (actual arrival tracking)
✅ Status Display (NSTATUS shown in timeline)
✅ Delay Detection (from email bodies)
✅ Source Tracking (📧 email vs 📊 Excel)
✅ Conflict Detection (email vs Excel discrepancies)
✅ Confidence Scoring (how certain we are about updates)

EXISTING FEATURES:
✅ Calendar Integration
✅ ETA Countdown
✅ Auto-Checklists with color coding
✅ Multi-Jetty Timeline
✅ Weather & Pilot Status
✅ Vessel Tracking
✅ Smart Email Processing
✅ Teams Integration
"""

import win32com.client
from datetime import datetime, timedelta
import requests
import re
import os
import time
import json
import traceback
import urllib3
import schedule

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# CONFIGURATION
# =========================================================

TEAMS_WEBHOOK_URL = "https://shellcorp.webhook.office.com/webhookb2/ccef7b6d-43a9-474c-80fb-5d777df8eabb@db1e96a8-a3da-442a-930b-235cac24cd5c/IncomingWebhook/0052852c3fbd41249fd1124de01567b5/03fc38bc-e1e1-4ccd-97c2-48d2590d89c6/V2SHLXrAyc8GCBCRFpHGm5tYGFMuicvW0zMcICZx8jN9I1"

PROXIES = {
    "http": "http://zproxy-global.shell.com:80",
    "https": "http://zproxy-global.shell.com:80"
}

WEATHER_API_KEY = "d0b00430f81bec691fdc8c46101afa73"
PORT_LAT = 51.9244
PORT_LON = 4.4777

WIND_WARN_THRESHOLD = 25
WIND_CRITICAL_THRESHOLD = 35
VISIBILITY_WARN = 1
TEMP_FREEZING = 0

WORK_HOURS_START = 7
WORK_HOURS_END = 18
RUN_INTERVAL_HOURS = 1
ETA_ALERT_HOURS = 24

PILOT_STATUS_FILE = "pilot_status.json"
PILOT_EMAIL = "hcc@portofrotterdam.com"
PILOT_KEYWORDS = ["pilot service", "pilot services", "pilotage", "pin rotterdam", "port information notice"]

CHECKLIST_FILE = "vessel_checklists.json"
TIMELINE_FILE = "jetty_timeline.json"
STATE_FILE = "processed_state.json"
LOG_FILE = "agent_log.txt"
WEEKLY_STATS_FILE = "weekly_stats.json"

JETTY_CONFIG = {
    "ST2": {"name": "Single Jetty 2", "min_length": 99, "max_length": 190},
    "ST3": {"name": "Single Jetty 3", "min_length": 99, "max_length": 155},
    "ST4": {"name": "Single Jetty 4", "min_length": 85, "max_length": 185},
    "ST5": {"name": "Single Jetty 5", "min_length": 85, "max_length": 111},
    "ST15": {"name": "Single Jetty 15", "min_length": 60, "max_length": 90},
    "ST16": {"name": "Single Jetty 16", "min_length": 60, "max_length": 90},
    "ST17": {"name": "Single Jetty 17", "min_length": 85, "max_length": 190},
    "ST18": {"name": "Single Jetty 18", "min_length": 85, "max_length": 185},
    "ST35": {"name": "Single Jetty 35", "min_length": 85, "max_length": 185},
    "ST35A": {"name": "Single Jetty 35A", "min_length": 85, "max_length": 185}
}

KNOWN_VESSELS = {
    "TEMPEST": "9424754", "SEFARINA": "9715701", "CHEMICAL LUNA": "9521423",
    "SFL BONAIRE": "9919773", "XING TONG KAI YUAN": "9640126", "VOYAGER": "02332403",
    "KENTERING": "02211189", "LEONARDO": "07001724", "STOLT MERWEDE": "9232490",
    "BARCELONA": "9233647", "BITHAV": "9999998", "VICTROL": "9999999", "UNIGAS II": "02340295",
    "BAYAMO": "9655004", "UNIGAS III": "EN02340282", "UNIGAS I": "EN02340295"
}

AGENT_EMAILS = ["wilhelmsen.com", "lbhnetherlands.com", "chemship.com", "vertomcory.com", "iss-shipping.com"]

CATEGORY_KEYWORDS = {
    "AGENT": [
        "voy", "voyage", "eta", "bunkers", "pilots", "laytime", "PORTBASE",
        "eta", "etb", "etd", "arrival", "departure", "crew",
        "port call", "laycan", "advised eta"
    ],
    "TERMINAL": [
        "berth", "jetty", "st18", "st35", "st17", "st4", "st18",
        "mooring", "terminal", "loading rate", "shore"
    ],
    "SURVEYOR": [
        "survey", "sgs", "bureau veritas", "intertek", "saybolt",
        "COA", "ullage", "sample"
    ],
    "NOMINATION": [
        "stem", " grade", "bill of lading", "b/l",
        "AMENDMENT", "BC FULL NOM"
    ],
    "LOADING_MASTER": [
        "loading plan", "discharge plan", "cargo plan",
        "tank allocation", "loading sequence"
    ],
    "OPERATIONS": [
        "schedule", "planning", "update", "delay", "waiting",
        "coordination", "meeting"
    ]
}

CATEGORY_ORDER = ["HIGH PRIORITY", "TERMINAL", "AGENT", "SURVEYOR", "LOADING_MASTER", "NOMINATION", "team lead"]
DELAY_KEYWORDS = ["awaiting", "delay", "delayed", "maintenance", "hold", "weather"]

# =========================================================
# NEW: EMAIL-TO-CHECKLIST KEYWORD MAPPINGS
# =========================================================

CHECKLIST_KEYWORDS = {
    'Pilot booking confirmed': {
        'positive': ['pilot ordered', 'pilot on board', 'incoming pilot ordered', 'pilot confirmed', 'pilotage arranged', 'pilot booked'],
        'negative': ['awaiting pilot', 'pilot tbc', 'pilot pending', 'pilot not', 'no pilot']
    },
    'Berth availability confirmed': {
        'positive': ['all fast', 'first line ashore', 'berth confirmed', 'gangway down', 'vessel moored', 'berth allocated', 'jetty confirmed'],
        'negative': ['awaiting berth availability', 'awaiting berth', 'berth tbc', 'waiting for berth', 'no berth']
    },
    'Agent notified': {
        'positive': ['notice of readiness tendered', 'notice of readiness received', 'nor tendered', 'nor received', 'nor submitted'],
        'negative': []
    },
    'Surveyor booked': {
        'positive': ['surveyor on board', 'surveyor confirmed', 'samples taken', 'calculations completed', 'sgs confirmed', 'intertek confirmed', 'surveyor will attend'],
        'negative': ['surveyor tbc', 'awaiting surveyor', 'no surveyor', 'surveyor not']
    },
    'Loading plan approved': {
        'positive': ['cargo operations resumed', 'commence discharging', 'commence loading', 'operations commenced', 'cargo operations started', 'loading commenced', 'discharge commenced'],
        'negative': ['cargo operations suspended', 'operations on hold', 'operations suspended', 'loading suspended']
    },
    'Mooring crew ready': {
        'positive': ['all fast', 'first line ashore', 'vessel moored', 'gangway down', 'mooring complete'],
        'negative': []
    }
}

DELAY_INDICATORS = [
    'cargo operations suspended',
    'delay - waiting',
    'operations suspended',
    'expect to commence',
    'revised eta',
    'postponed',
    'on hold',
    'due insufficient',
    'awaiting berth',
    'waiting for'
]

# =========================================================
# LOGGING
# =========================================================

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
        print(msg)
    except:
        print(msg)

def log_exception(e):
    log(f"ERROR: {str(e)}")
    try:
        log(traceback.format_exc())
    except:
        pass

def retry(max_attempts=3, delay=1.0):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as exc:
                    log(f"Retry {attempt}/{max_attempts} for {func.__name__}: {exc}")
                    time.sleep(delay)
            return None
        return wrapper
    return decorator

# =========================================================
# STATE MANAGEMENT
# =========================================================

def load_processed_ids():
    if not os.path.exists(STATE_FILE):
        return set()
    try:
        with open(STATE_FILE, "r") as f:
            return set(json.load(f))
    except:
        return set()

def save_processed_ids(ids):
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(list(ids), f)
    except:
        pass

def load_timeline():
    if not os.path.exists(TIMELINE_FILE):
        return {"vessels": [], "maintenance": []}
    try:
        with open(TIMELINE_FILE, "r") as f:
            return json.load(f)
    except:
        return {"vessels": [], "maintenance": []}

def save_timeline(timeline):
    try:
        with open(TIMELINE_FILE, "w") as f:
            json.dump(timeline, f, indent=2)
    except:
        pass

def load_checklists():
    if not os.path.exists(CHECKLIST_FILE):
        return {}
    try:
        with open(CHECKLIST_FILE, "r") as f:
            return json.load(f)
    except:
        return {}

def save_checklists(checklists):
    try:
        with open(CHECKLIST_FILE, "w") as f:
            json.dump(checklists, f, indent=2)
    except:
        pass

def load_pilot_status():
    if not os.path.exists(PILOT_STATUS_FILE):
        return None
    try:
        with open(PILOT_STATUS_FILE, "r") as f:
            return json.load(f)
    except:
        return None

def save_pilot_status(status):
    try:
        with open(PILOT_STATUS_FILE, "w") as f:
            json.dump(status, f, indent=2)
    except:
        pass

# =========================================================
# VESSEL FUNCTIONS
# =========================================================

def detect_identifier_type(identifier):
    if not identifier:
        return None
    id_str = str(identifier)
    return 'ENI' if len(id_str) == 8 else 'IMO' if len(id_str) == 7 else None

def build_vessel_url(vessel_name, identifier, identifier_type='IMO'):
    try:
        if identifier and identifier_type == 'IMO':
            return f"https://www.vesselfinder.com/?imo={identifier}"
        elif identifier and identifier_type == 'ENI':
            return f"https://www.marinetraffic.com/en/ais/index/search/all?keyword={identifier}"
        elif vessel_name:
            return f"https://www.vesselfinder.com/vessels?name={vessel_name.replace(' ', '%20')}"
        return "https://www.vesselfinder.com"
    except:
        return "https://www.vesselfinder.com"

def extract_vessel_names(text):
    try:
        found = []
        text_upper = text.upper()
        for vessel in KNOWN_VESSELS.keys():
            if vessel in text_upper and vessel not in found:
                found.append(vessel)
        return found[:3]
    except:
        return []

def collect_vessel_info(emails):
    vessels_info = {}
    for email in emails:
        for vessel_name in email.get('vessels', []):
            if vessel_name not in vessels_info:
                vessel_id = KNOWN_VESSELS.get(vessel_name)
                id_type = detect_identifier_type(vessel_id)
                vessel_url = build_vessel_url(vessel_name, vessel_id, id_type)
                vessels_info[vessel_name] = {
                    'identifier': vessel_id,
                    'identifier_type': id_type,
                    'emails': [],
                    'categories': set(),
                    'vessel_url': vessel_url
                }
            vessels_info[vessel_name]['emails'].append(email)
            vessels_info[vessel_name]['categories'].add(email['category'])
    return vessels_info

# =========================================================
# ETA COUNTDOWN
# =========================================================

def get_eta_countdown(eta_date_str, anchored_date=None):
    """Calculate countdown to ETA with color coding - only ARRIVED if anchored_date exists"""
    try:
        # Check if actually arrived (has anchored_date in past)
        if anchored_date and str(anchored_date).strip():
            try:
                anchored = datetime.fromisoformat(str(anchored_date).strip())
                if anchored <= datetime.now():
                    return "⚓ ARRIVED", "Good"
            except:
                pass
        
        # Otherwise calculate countdown
        eta = datetime.fromisoformat(eta_date_str)
        now = datetime.now()
        delta = eta - now
        hours = delta.total_seconds() / 3600

        if hours < 0:
            return "⏰ Overdue", "Warning"
        elif hours < 6:
            return f"🚨 ARRIVING IN {int(hours)}h", "Attention"
        elif hours < 24:
            return f"⏰ Arriving in {int(hours)}h", "Warning"
        elif hours < 48:
            return f"📅 Tomorrow ({int(hours)}h)", "Good"
        else:
            days = int(hours / 24)
            return f"📅 In {days} days", "Default"
    except:
        return "📅 Scheduled", "Default"
# =========================================================
# NEW: SMART EMAIL PARSING FOR CHECKLISTS
# =========================================================

def parse_email_for_checklist_updates(email_body, email_subject, vessel_name, sender_name):
    """
    Intelligent email parsing to detect checklist item completions
    Returns: {
        'updates': [{task, confidence, source, reason}],
        'delays': [{message, source}],
        'conflicts': [{message}]
    }
    """
    try:
        result = {
            'updates': [],
            'delays': [],
            'conflicts': []
        }
        
        combined_text = f"{email_subject} {email_body}".lower()
        
        # Check for delay indicators
        for delay_keyword in DELAY_INDICATORS:
            if delay_keyword in combined_text:
                result['delays'].append({
                    'message': f"Delay indicator: '{delay_keyword}'",
                    'source': f"Email from {sender_name}"
                })
        
        # Check each checklist item
        for task_name, keywords in CHECKLIST_KEYWORDS.items():
            positive_matches = []
            negative_matches = []
            
            # Check positive keywords
            for keyword in keywords['positive']:
                if keyword in combined_text:
                    positive_matches.append(keyword)
            
            # Check negative keywords
            for keyword in keywords['negative']:
                if keyword in combined_text:
                    negative_matches.append(keyword)
            
            # Determine if we should update
            if positive_matches and not negative_matches:
                confidence = min(len(positive_matches) * 30 + 40, 95)
                result['updates'].append({
                    'task': task_name,
                    'confidence': confidence,
                    'source': f"📧 Email from {sender_name}",
                    'reason': f"Found: {', '.join(positive_matches[:2])}"
                })
            elif negative_matches:
                result['conflicts'].append({
                    'message': f"{task_name}: Negative indicator found ('{negative_matches[0]}')"
                })
        
        return result
        
    except Exception as e:
        log(f"Error parsing email: {e}")
        return {'updates': [], 'delays': [], 'conflicts': []}

def update_checklists_from_emails(checklists, emails):
    """
    Update checklists based on email content
    Returns updated checklists and summary of changes
    """
    try:
        updates_made = 0
        delays_detected = []
        
        log("📧 Analyzing emails for checklist updates...")
        
        for email in emails:
            vessels = email.get('vessels', [])
            
            for vessel_name in vessels:
                if vessel_name not in checklists:
                    continue
                
                # Parse email for updates
                parse_result = parse_email_for_checklist_updates(
                    email.get('body', ''),
                    email.get('subject', ''),
                    vessel_name,
                    email.get('sender_name', 'Unknown')
                )
                
                # Apply updates
                checklist = checklists[vessel_name]
                
                for update in parse_result['updates']:
                    for item in checklist.get('items', []):
                        if item['task'] == update['task'] and item.get('status') != 'COMPLETED':
                            item['status'] = 'COMPLETED'
                            item['completed_by'] = update['source']
                            item['completed_at'] = datetime.now().isoformat()
                            item['confidence'] = update['confidence']
                            item['reason'] = update['reason']
                            updates_made += 1
                            log(f"   ✅ {vessel_name}: {item['task']} (confidence: {update['confidence']}%)")
                
                # Track delays
                if parse_result['delays']:
                    delays_detected.append({
                        'vessel': vessel_name,
                        'delays': parse_result['delays']
                    })
        
        if updates_made > 0:
            log(f"✅ Email analysis: {updates_made} checklist items updated")
            save_checklists(checklists)
        else:
            log("   No email-based updates needed")
        
        return checklists, delays_detected
        
    except Exception as e:
        log(f"Error updating from emails: {e}")
        log_exception(e)
        return checklists, []

# =========================================================
# TIMELINE WITH ANCHORED DATE & STATUS
# =========================================================

def detect_conflicts(timeline):
    conflicts = []
    try:
        by_jetty = {}
        for v in sorted([v for v in timeline['vessels'] if v.get('eta')], key=lambda x: x['eta']):
            jetty = v.get('jetty', 'TBD')
            if jetty not in by_jetty:
                by_jetty[jetty] = []
            by_jetty[jetty].append(v)

        for jetty, vessels in by_jetty.items():
            for i in range(len(vessels) - 1):
                v1, v2 = vessels[i], vessels[i + 1]
                if v1.get('etd') and v2.get('eta'):
                    gap = (datetime.fromisoformat(v2['eta']) - datetime.fromisoformat(v1['etd'])).total_seconds() / 3600
                    if gap < 0:
                        conflicts.append({
                            'severity': 'CRITICAL',
                            'message': f"⛔ OVERLAP at {jetty}: {v1['name']} & {v2['name']}"
                        })
                    elif gap < 2:
                        conflicts.append({
                            'severity': 'WARNING',
                            'message': f"⚠️ TIGHT: {gap:.1f}h gap at {jetty} ({v1['name']} → {v2['name']})"
                        })
        return conflicts
    except:
        return []

def build_timeline_visualization(timeline, days=7):
    """Timeline with ANCHORED_DATE logic and STATUS display"""
    try:
        now = datetime.now()
        today = now.date()
        end_day = (now + timedelta(days=days)).date()

        visible = []
        for v in timeline.get('vessels', []):
            if not v.get('eta'):
                continue

            try:
                eta_day = datetime.fromisoformat(v['eta']).date()
                
                # Check if vessel has arrived (based on anchored_date)
                has_arrived = False
                if v.get('anchored_date'):
                    try:
                        anchored = datetime.fromisoformat(v['anchored_date'])
                        if anchored <= now:
                            has_arrived = True
                    except:
                        pass
                
                # Include if arrived or arriving within window
                if has_arrived or (today <= eta_day <= end_day):
                    visible.append(v)
                    
            except:
                continue

        if not visible:
            return None, []

        visible.sort(key=lambda x: (x.get('jetty', 'ZZZ'), x.get('eta', '')))

        text = ""
        vessel_actions = []
        current_jetty = None

        for v in visible:
            vessel_name = v['name']
            jetty = v.get('jetty', 'TBD')
            cargo = v.get('cargo', 'TBC')[:35]
            agent = v.get('agent', '')
            status_desc = v.get('status_desc', '')

            if jetty != current_jetty:
                if current_jetty is not None:
                    text += "\n\n"
                text += f"**{jetty}** "
                current_jetty = jetty

            vessel_id = KNOWN_VESSELS.get(vessel_name) or v.get('imo')
            id_type = detect_identifier_type(vessel_id)
            vessel_url = build_vessel_url(vessel_name, vessel_id, id_type)

            vessel_actions.append({
                "type": "Action.OpenUrl",
                "title": f"📍 {vessel_name[:12]}",
                "url": vessel_url
            })

            icon = "⛵" if id_type == 'ENI' else "🚢"

            # Determine status based on anchored_date
            has_arrived = False
            if v.get('anchored_date'):
                try:
                    anchored = datetime.fromisoformat(v['anchored_date'])
                    if anchored <= now:
                        has_arrived = True
                except:
                    pass
            
            if has_arrived:
                status_text = "⚓ ARRIVED"
                status_icon = "✅"
            else:
                countdown_text, _ = get_eta_countdown(v['eta'])
                status_text = countdown_text
                status_icon = "⏳"

            try:
                eta_str = datetime.fromisoformat(v['eta']).strftime("%b %d")
                etd_str = datetime.fromisoformat(v['etd']).strftime("%b %d") if v.get('etd') else "TBC"
            except:
                eta_str = "TBC"
                etd_str = "TBC"

            # Build vessel entry with STATUS
            text += f"{icon} **{vessel_name}** {status_icon} | {cargo} | 📊 {status_desc} | {status_text} 📅 {eta_str} → {etd_str}"
            
            if agent:
                text += f" | 👤 {agent}"
            
            surveyor = v.get('surveyor', '')
            if surveyor and surveyor not in ['NONE', 'NIET VAN TOEPASSING', '']:
                text += f" | 🔍 {surveyor}"
            
            text += " "

        return text, vessel_actions

    except Exception as e:
        log(f"Timeline error: {e}")
        log_exception(e)
        return None, []

# =========================================================
# CHECKLISTS WITH SMART AUTO-COMPLETION
# =========================================================

def create_arrival_checklist(vessel_name, eta, jetty=None):
    try:
        vessel_id = KNOWN_VESSELS.get(vessel_name)
        is_barge = vessel_id and len(str(vessel_id)) == 8

        items = [
            {'task': 'Pilot booking confirmed', 'deadline': '24h', 'status': 'PENDING', 'critical': True},
            {'task': 'Berth availability confirmed', 'deadline': '24h', 'status': 'PENDING', 'critical': True},
            {'task': 'Agent notified', 'deadline': '24h', 'status': 'PENDING', 'critical': True},
        ]

        if not is_barge:
            items.append({'task': 'Surveyor booked', 'deadline': '24h', 'status': 'PENDING', 'critical': True})

        items.extend([
            {'task': 'Loading plan approved', 'deadline': '12h', 'status': 'PENDING', 'critical': True},
            {'task': 'Mooring crew ready', 'deadline': 'Arrival', 'status': 'PENDING', 'critical': True}
        ])

        return {
            'vessel': vessel_name,
            'eta': eta['date'] if isinstance(eta, dict) else eta,
            'jetty': jetty or 'TBD',
            'created': datetime.now().isoformat(),
            'items': items
        }
    except:
        return None

def update_checklists(timeline):
    """Smart checklist updates from Excel data"""
    try:
        checklists = load_checklists()
        now = datetime.now()
        updates_made = 0

        log("📋 Updating checklists from Excel...")

        for vessel in timeline.get('vessels', []):
            vessel_name = vessel['name']
            eta = vessel.get('eta')
            jetty = vessel.get('jetty')

            status_desc = vessel.get('status_desc', '').lower()
            ship_inspector = vessel.get('ship_inspector', 'NONE')

            if not eta:
                continue

            try:
                eta_date = datetime.fromisoformat(eta)
                hours_until = (eta_date - now).total_seconds() / 3600

                if -24 < hours_until < 72:
                    if vessel_name not in checklists:
                        new_checklist = create_arrival_checklist(vessel_name, {'date': eta}, jetty)
                        if new_checklist:
                            checklists[vessel_name] = new_checklist
                            log(f"   📋 Created checklist: {vessel_name} (ETA in {hours_until:.1f}h)")

                    if vessel_name in checklists:
                        checklist = checklists[vessel_name]

                        for item in checklist.get('items', []):
                            task_lower = item['task'].lower()

                            if item.get('status') == 'COMPLETED':
                                continue

                            if "released to operations" in status_desc:
                                if any(kw in task_lower for kw in ['agent notified', 'berth availability', 'loading plan']):
                                    item['status'] = 'COMPLETED'
                                    item['completed_by'] = '📊 AUTO: Excel'
                                    item['completed_at'] = now.isoformat()
                                    updates_made += 1
                                    log(f"   ✅ {vessel_name}: {item['task']}")

                            if 'surveyor' in task_lower:
                                if ship_inspector not in ['NONE', '', 'NA', 'TBA', 'NIET VAN TOEPASSING']:
                                    item['status'] = 'COMPLETED'
                                    item['completed_by'] = f'📊 AUTO: {ship_inspector}'
                                    item['completed_at'] = now.isoformat()
                                    updates_made += 1
                                    log(f"   ✅ {vessel_name}: Surveyor ({ship_inspector})")

            except Exception as e:
                log(f"   ⚠️ Error processing {vessel_name}: {e}")
                continue

        if updates_made > 0:
            log(f"✅ Excel-based: {updates_made} checklist items completed")
        else:
            log("   No Excel-based updates needed")

        save_checklists(checklists)
        return checklists

    except Exception as e:
        log(f"Checklist update error: {e}")
        log_exception(e)
        return load_checklists()

def cleanup_old_checklists(checklists, timeline):
    """Remove checklists for vessels no longer in timeline or >72h away"""
    try:
        now = datetime.now()
        timeline_vessels = [v['name'] for v in timeline.get('vessels', [])]

        to_remove = []

        for vessel_name, checklist in checklists.items():
            if vessel_name not in timeline_vessels:
                to_remove.append(vessel_name)
                continue

            eta = checklist.get('eta')
            if eta:
                try:
                    eta_date = datetime.fromisoformat(eta)
                    hours_until = (eta_date - now).total_seconds() / 3600

                    if hours_until < -24 or hours_until > 72:
                        to_remove.append(vessel_name)
                except:
                    pass

        for vessel_name in to_remove:
            del checklists[vessel_name]
            log(f"🗑️ Removed old checklist: {vessel_name}")

        if to_remove:
            save_checklists(checklists)
            log(f"✅ Cleaned up {len(to_remove)} old checklists")

        return checklists

    except Exception as e:
        log(f"Cleanup error: {e}")
        return checklists

def get_checklist_summary(checklists):
    """Generate checklist summary for display"""
    summary = {'total': 0, 'at_risk': []}
    try:
        now = datetime.now()

        for vessel_name, checklist in checklists.items():
            eta = checklist.get('eta')
            if not eta:
                continue

            try:
                eta_date = datetime.fromisoformat(eta)
                hours_until = (eta_date - now).total_seconds() / 3600
            except:
                continue

            if hours_until < -24 or hours_until > 72:
                continue

            summary['total'] += 1

            pending_critical = []
            completed = 0

            for item in checklist.get('items', []):
                if item.get('status') == 'PENDING':
                    if item.get('critical'):
                        pending_critical.append(item['task'])
                else:
                    completed += 1

            total_items = len(checklist.get('items', []))
            completion_pct = int((completed / total_items * 100)) if total_items > 0 else 0

            summary['at_risk'].append({
                'vessel': vessel_name,
                'eta': eta,
                'hours_until': round(hours_until, 1),
                'jetty': checklist.get('jetty', 'TBD'),
                'pending_critical': pending_critical,
                'completed': completed,
                'total': total_items,
                'completion_pct': completion_pct
            })

        summary['at_risk'].sort(key=lambda x: x['hours_until'])
        return summary

    except Exception as e:
        log(f"Summary error: {e}")
        return summary

# =========================================================
# WEATHER & PILOT
# =========================================================

def get_wind_direction(degrees):
    dirs = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
    return dirs[int((degrees + 11.25) / 22.5) % 16]

@retry(max_attempts=2, delay=2.0)
def get_weather_conditions():
    try:
        response = requests.get(
            "https://api.openweathermap.org/data/2.5/weather",
            params={"lat": PORT_LAT, "lon": PORT_LON, "appid": WEATHER_API_KEY, "units": "metric"},
            proxies=PROXIES, timeout=10, verify=False
        )
        response.raise_for_status()
        data = response.json()

        temp = data['main']['temp']
        wind_kt = data['wind']['speed'] * 1.944
        wind_dir = get_wind_direction(data['wind'].get('deg', 0))

        if wind_kt >= WIND_CRITICAL_THRESHOLD:
            ops_safe, ops_status, wind_status = False, "⛔ OPERATIONS UNSAFE", "🔴 CRITICAL"
        elif wind_kt >= WIND_WARN_THRESHOLD:
            ops_safe, ops_status, wind_status = False, "⚠️ MONITOR OPERATIONS", "🟡 CAUTION"
        else:
            ops_safe, ops_status, wind_status = True, "✅ OPERATIONS NORMAL", "🟢 NORMAL"

        return {
            "temperature": round(temp, 1),
            "feels_like": round(data['main']['feels_like'], 1),
            "wind_speed": round(wind_kt, 1),
            "wind_direction": wind_dir,
            "wind_status": wind_status,
            "visibility": round(data.get('visibility', 10000) / 1000, 1),
            "conditions": data['weather'][0]['description'].title(),
            "safe_for_operations": ops_safe,
            "operational_status": ops_status,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    except:
        return None

def is_pilot_service_email(sender, subject, body):
    try:
        if PILOT_EMAIL.lower() in sender.lower():
            return True
        combined = f"{subject} {body}".lower()
        return any(kw in combined for kw in PILOT_KEYWORDS)
    except:
        return False

def parse_pilot_service_status(body, subject):
    """Parse pilot status with actual subject line"""
    try:
        combined = f"{subject} {body}".lower()

        clean_subject = subject
        if "pin rotterdam" in subject.lower():
            clean_subject = subject.split("-", 1)[-1].strip() if "-" in subject else subject

        if "normal" in combined or "resumed" in combined or "lifted" in combined:
            return {
                "status": "NORMAL",
                "status_emoji": "✅",
                "status_text": clean_subject,
                "status_message": clean_subject,
                "color": "Good",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "operational": True
            }
        elif "suspended" in combined or "restricted" in combined or "closed" in combined:
            return {
                "status": "SUSPENDED",
                "status_emoji": "🔴",
                "status_text": clean_subject,
                "status_message": clean_subject,
                "color": "Attention",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "operational": False
            }
        else:
            return {
                "status": "UPDATE",
                "status_emoji": "ℹ️",
                "status_text": clean_subject,
                "status_message": clean_subject,
                "color": "Default",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "operational": True
            }
    except:
        return None

# =========================================================
# EMAIL PROCESSING
# =========================================================

def compute_delay_risk(text):
    score = sum(1 for word in DELAY_KEYWORDS if word in text.lower())
    return "HIGH" if score >= 5 else "MEDIUM" if score >= 3 else "LOW" if score >= 1 else "NONE"

def extract_summary(body):
    """Intelligent email summary extraction"""
    try:
        if not body or len(body) < 10:
            return "(No content)"

        body = body.strip()
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]

        clean_lines = []
        for line in lines:
            if any(x in line.lower() for x in ["best regards", "kind regards", "yours sincerely", "thanks,", "regards,", "sent from my", "this email and any attachments"]):
                break
            if len(line) < 5:
                continue
            if not any(c.isalnum() for c in line):
                continue
            clean_lines.append(line)

        if not clean_lines:
            return "(No meaningful content)"

        questions = [ln for ln in clean_lines if "?" in ln]
        if questions:
            return " ".join(questions[:2])

        action_keywords = ["please", "need", "require", "confirm", "advise", "update", "inform", "request", "check", "arrange", "book", "schedule"]
        action_lines = [ln for ln in clean_lines if any(kw in ln.lower() for kw in action_keywords)]
        if action_lines:
            return " ".join(action_lines[:2])

        info_keywords = ["vessel", "cargo", "eta", "etd", "laycan", "berth", "jetty", "loading", "discharge"]
        info_lines = [ln for ln in clean_lines if any(kw in ln.lower() for kw in info_keywords)]
        if info_lines:
            return " ".join(info_lines[:2])

        substantial = [ln for ln in clean_lines if len(ln) > 20]
        if substantial:
            return " ".join(substantial[:2])

        return " ".join(clean_lines[:2])

    except Exception as e:
        log(f"Error extracting summary: {e}")
        return "(Summary unavailable)"

def categorize_email(email):
    """Categorize email from SUBJECT + SENDER only"""
    subject = email.get('subject', '').lower()
    sender = email.get('sender_email', '').lower()
    sender_name = email.get('sender_name', '').lower()

    text = f"{subject} {sender} {sender_name}"

    if any(domain in sender for domain in AGENT_EMAILS):
        return "AGENT"

    if any(kw in subject for kw in ["urgent", "asap", "critical", "immediate", "priority"]):
        return "HIGH PRIORITY"

    category_scores = {}
    for cat, keywords in CATEGORY_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in subject)
        if score > 0:
            category_scores[cat] = score

    if category_scores:
        best_category = max(category_scores, key=category_scores.get)
        return best_category

    vessels = email.get('vessels', [])
    if vessels and len(vessels) > 0:
        if any(word in subject for word in ["voy", "voyage", "eta", "etd", "laycan", "nomination"]):
            return "AGENT"
        return "OPERATIONS"

    return "GENERAL"

def calculate_urgency_score(email):
    score = 0
    text = f"{email.get('subject', '')} {email.get('body', '')}".lower()

    if any(w in text for w in ["urgent", "asap", "critical"]):
        score += 30
    risk = email.get('delay_risk', 'NONE')
    if risk == 'HIGH':
        score += 25
    elif risk == 'MEDIUM':
        score += 15

    return min(score, 100)

def get_urgency_emoji(score):
    return "🔴" if score >= 70 else "🟡" if score >= 50 else "🟢" if score >= 30 else "⚪"

def generate_smart_reply(email, weather=None):
    """Generate professional reply - ONLY FOR URGENT EMAILS"""
    try:
        sender = email.get('sender_name', 'Sir/Madam').split()[0]
        subject = email.get('subject', 'your message')
        vessels = email.get('vessels', [])

        reply = f"Dear {sender},\n\n"
        reply += f"Thank you for your urgent message regarding {subject.lower()}.\n\n"

        if vessels:
            reply += f"Vessel: {vessels[0]}\n"

        if weather:
            reply += f"Current conditions: {weather['operational_status']}\n"

        reply += "\nWe are giving this matter immediate attention and will update you shortly.\n\n"
        reply += "Best regards,\nShell Rotterdam Chemical Terminal"

        return reply
    except:
        return None

# =========================================================
# CALENDAR
# =========================================================

@retry()
def fetch_calendar():
    """Bulletproof calendar fetch"""
    try:
        log("📅 Fetching calendar...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        cal = ns.GetDefaultFolder(9)

        now = datetime.now()
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        today_end = now.replace(hour=23, minute=59, second=59, microsecond=0)

        start_str = today_start.strftime("%m/%d/%Y %H:%M %p")
        end_str = today_end.strftime("%m/%d/%Y %H:%M %p")

        restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"

        log(f"📅 Searching for events on: {now.strftime('%Y-%m-%d')}")

        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        restricted_items = items.Restrict(restriction)

        events = []

        for item in restricted_items:
            try:
                if item.Class != 26:
                    continue

                subject = item.Subject or "(No Subject)"

                start = item.Start
                end = item.End

                if item.AllDayEvent:
                    start_time = "All day"
                    end_time = ""
                else:
                    if hasattr(start, 'strftime'):
                        start_time = start.strftime("%H:%M")
                        end_time = end.strftime("%H:%M")
                    else:
                        start_time = "TBD"
                        end_time = ""

                location = item.Location or "Not specified"

                is_teams = False
                if "Microsoft Teams" in location or "teams.microsoft.com" in str(item.Body):
                    is_teams = True

                try:
                    organizer = item.Organizer
                except:
                    organizer = "Unknown"

                events.append({
                    "subject": subject,
                    "start_time": start_time,
                    "end_time": end_time,
                    "location": location,
                    "organizer": organizer,
                    "is_teams": is_teams
                })

                log(f"   ✅ {subject} at {start_time}")

            except Exception as e:
                log(f"   ⚠️ Error processing appointment: {e}")
                continue

        log(f"📅 Found {len(events)} calendar events")

        return events

    except Exception as e:
        log(f"❌ Calendar error: {e}")
        log_exception(e)
        return []

# =========================================================
# FETCH EMAILS
# =========================================================

@retry()
def fetch_emails():
    log("📧 Fetching emails...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    msgs = inbox.Items
    msgs.Sort("[ReceivedTime]", True)

    processed_ids = load_processed_ids()
    new_emails = []

    for msg in msgs:
        try:
            if not msg.UnRead or msg.EntryID in processed_ids:
                continue

            sender_email = msg.SenderEmailAddress or "Unknown"
            subject = msg.Subject or "(No Subject)"
            body = msg.Body or ""

            if is_pilot_service_email(sender_email, subject, body):
                log(f"📍 Pilot email: {subject}")
                status = parse_pilot_service_status(body, subject)
                if status:
                    save_pilot_status(status)
                msg.UnRead = False
                processed_ids.add(msg.EntryID)
                continue

            vessels = extract_vessel_names(f"{subject}\n{body}")
            
            email_data = {
                "sender_name": msg.SenderName or "Unknown",
                "sender_email": sender_email,
                "subject": subject,
                "body": body[:200].replace('"', "'").replace('\n', ' ').replace('\r', ''),
                "smart_summary": extract_summary(body).replace('"', "'").replace('\n', ' '),
                "received_time": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                "entry_id": msg.EntryID,
                "vessels": vessels,
                "delay_risk": compute_delay_risk(body),
                "category": "",
                "urgency_score": 0
            }
            
            email_data["category"] = categorize_email(email_data)
            email_data["urgency_score"] = calculate_urgency_score(email_data)
            
            # ✅ FLAG EMAIL BEFORE MARKING AS READ!
            try:
                # Apply category
                msg.Categories = email_data["category"]
                
                # Flag if urgent
                if email_data["urgency_score"] >= 70:
                    msg.FlagRequest = "Urgent - Requires immediate attention"
                    msg.FlagStatus = 2  # olFlagMarked (red flag)
                    msg.Importance = 2  # olImportanceHigh
                    log(f"   🚩 Flagged urgent: {subject[:50]}")
                
                msg.Save()  # Save changes
            except Exception as e:
                log(f"   ⚠️ Could not flag email: {e}")
            
            new_emails.append(email_data)
            msg.UnRead = False
            processed_ids.add(msg.EntryID)

        except Exception as e:
            log(f"Error processing email: {e}")
            continue

    save_processed_ids(processed_ids)
    log(f"✅ Fetched {len(new_emails)} emails")
    return new_emails

# =========================================================
# TEAMS SUMMARY WITH ALL NEW FEATURES
# =========================================================

def send_summary_to_teams(emails, events, weather, vessels_info, pilot_status, timeline, conflicts, checklist_summary, checklists, delays_detected):
    try:
        log("📤 Building Teams card...")

        card_body = []

        card_body.append({
            "type": "TextBlock",
            "text": "📋 Daily Jetty Planning & Email Summary",
            "size": "ExtraLarge",
            "weight": "Bolder",
            "color": "Accent"
        })

        if weather:
            card_body.append({"type": "TextBlock", "text": "🌤 Live Port Conditions", "size": "Large", "weight": "Bolder", "separator": True, "spacing": "Large"})
            card_body.append({"type": "TextBlock", "text": weather['operational_status'], "weight": "Bolder", "color": "Attention" if not weather['safe_for_operations'] else "Good"})
            card_body.append({"type": "FactSet", "facts": [
                {"title": "🌡 Temp:", "value": f"{weather['temperature']}°C"},
                {"title": "💨 Wind:", "value": f"{weather['wind_speed']}kt {weather['wind_direction']} {weather['wind_status']}"},
                {"title": "🌊 Conditions:", "value": weather['conditions']}
            ]})

        if not pilot_status:
            pilot_status = load_pilot_status()
        if pilot_status:
            card_body.append({"type": "TextBlock", "text": "📍 Pilot Service", "size": "Large", "weight": "Bolder", "separator": True, "spacing": "Large"})
            card_body.append({"type": "TextBlock", "text": f"{pilot_status['status_emoji']} {pilot_status['status_text']}", "weight": "Bolder", "color": pilot_status['color']})

        timeline_viz, vessel_actions = build_timeline_visualization(timeline, days=7)

        if timeline_viz:
            card_body.append({
                "type": "TextBlock",
                "text": f"📊 Jetty Timeline (Next 7 Days)",
                "size": "Large",
                "weight": "Bolder",
                "separator": True,
                "spacing": "Large"
            })
            card_body.append({
                "type": "TextBlock",
                "text": timeline_viz,
                "wrap": True,
                "spacing": "Small"
            })

            if vessel_actions and len(vessel_actions) > 0:
                card_body.append({
                    "type": "TextBlock",
                    "text": "🔍 Quick Track:",
                    "weight": "Bolder",
                    "spacing": "Medium"
                })

                button_rows = []
                for i in range(0, len(vessel_actions), 3):
                    row = vessel_actions[i:i+3]
                    button_rows.append(row)

                for row in button_rows[:5]:
                    card_body.append({
                        "type": "ActionSet",
                        "actions": row,
                        "spacing": "Small"
                    })

        if conflicts:
            card_body.append({"type": "TextBlock", "text": f"⚠️ Conflicts ({len(conflicts)})", "size": "Large", "weight": "Bolder", "color": "Attention", "separator": True, "spacing": "Large"})
            for c in conflicts[:3]:
                card_body.append({"type": "TextBlock", "text": c['message'], "color": "Attention", "wrap": True})

        if delays_detected and len(delays_detected) > 0:
            card_body.append({
                "type": "TextBlock",
                "text": f"🚨 Delay Warnings Detected ({len(delays_detected)})",
                "size": "Large",
                "weight": "Bolder",
                "color": "Attention",
                "separator": True,
                "spacing": "Large"
            })
            for delay_info in delays_detected[:3]:
                vessel = delay_info['vessel']
                delay_messages = delay_info['delays']
                delay_text = f"**{vessel}**: " + "; ".join([d['message'] for d in delay_messages[:2]])
                card_body.append({
                    "type": "TextBlock",
                    "text": delay_text,
                    "wrap": True,
                    "color": "Attention"
                })

        if checklist_summary and checklist_summary.get('at_risk'):
            card_body.append({
                "type": "TextBlock",
                "text": f"📋 Arrival Checklists ({checklist_summary['total']} vessels <72h)",
                "size": "Large",
                "weight": "Bolder",
                "separator": True,
                "spacing": "Large"
            })

            for ck in checklist_summary['at_risk'][:5]:
                vessel_data = next((v for v in timeline.get('vessels', []) if v['name'] == ck['vessel']), None)
                anchored_date = vessel_data.get('anchored_date', '') if vessel_data else ''
                countdown, color = get_eta_countdown(ck['eta'], anchored_date)

                checklist_items = [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [{
                                    "type": "TextBlock",
                                    "text": f"🚢 **{ck['vessel']}** — {ck['jetty']}",
                                    "weight": "Bolder",
                                    "size": "Medium"
                                }]
                            },
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [{
                                    "type": "TextBlock",
                                    "text": countdown,
                                    "color": color,
                                    "weight": "Bolder",
                                    "horizontalAlignment": "Right"
                                }]
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": f"Progress: {ck['completed']}/{ck['total']} ({ck['completion_pct']}%)",
                        "spacing": "Small",
                        "isSubtle": True
                    }
                ]

                try:
                    vessel_checklist = checklists.get(ck['vessel'])
                    if vessel_checklist and vessel_checklist.get('items'):
                        completed_tasks = []
                        pending_tasks = []

                        for item in vessel_checklist['items']:
                            task_name = item['task']
                            status = item.get('status', 'PENDING')
                            completed_by = item.get('completed_by', '')

                            if status == 'COMPLETED':
                                task_display = f"✅ {task_name}"
                                if completed_by:
                                    task_display += f" ({completed_by})"
                                completed_tasks.append(task_display)
                            else:
                                pending_tasks.append(f"❌ {task_name}")

                        if completed_tasks:
                            checklist_items.append({
                                "type": "TextBlock",
                                "text": "**✅ Completed:**",
                                "weight": "Bolder",
                                "spacing": "Small",
                                "size": "Small",
                                "color": "Good"
                            })
                            checklist_items.append({
                                "type": "TextBlock",
                                "text": "\n".join(completed_tasks),
                                "wrap": True,
                                "spacing": "None",
                                "size": "Small",
                                "color": "Good"
                            })

                        if pending_tasks:
                            checklist_items.append({
                                "type": "TextBlock",
                                "text": "**❌ Pending:**",
                                "weight": "Bolder",
                                "spacing": "Small",
                                "size": "Small",
                                "color": "Attention"
                            })
                            checklist_items.append({
                                "type": "TextBlock",
                                "text": "\n".join(pending_tasks),
                                "wrap": True,
                                "spacing": "None",
                                "size": "Small",
                                "color": "Attention"
                            })
                except Exception as e:
                    log(f"Error displaying checklist for {ck['vessel']}: {e}")

                card_body.append({
                    "type": "Container",
                    "style": "emphasis",
                    "items": checklist_items,
                    "separator": True,
                    "spacing": "Small"
                })

        card_body.append({"type": "FactSet", "facts": [
            {"title": "📧 Emails:", "value": str(len(emails))},
            {"title": "🔴 Urgent:", "value": str(sum(1 for e in emails if e.get('urgency_score', 0) >= 70))},
            {"title": "🚢 Vessels:", "value": str(len(vessels_info))},
            {"title": "📊 Timeline:", "value": str(len(timeline.get('vessels', [])))}
        ], "separator": True})

        if len(emails) > 0:
            grouped = {}
            for e in emails:
                grouped.setdefault(e["category"], []).append(e)

            for cat in sorted(grouped.keys(), key=lambda c: CATEGORY_ORDER.index(c) if c in CATEGORY_ORDER else 999):
                items = grouped[cat]
                items.sort(key=lambda x: -x.get("urgency_score", 0))

                card_body.append({"type": "TextBlock", "text": f"🏷 {cat} ({len(items)})", "size": "Large", "weight": "Bolder", "separator": True, "spacing": "Large"})

                max_emails_per_category = 1 if len(emails) > 3 else 2
                for idx, e in enumerate(items[:max_emails_per_category], 1):
                    emoji = get_urgency_emoji(e.get('urgency_score', 0))
                    urgency = e.get('urgency_score', 0)

                    email_container = {"type": "Container", "style": "emphasis" if urgency >= 50 else "default", "items": [], "separator": True}
                    email_container["items"].append({"type": "TextBlock", "text": f"{idx}️⃣ {emoji} {e['subject']}", "weight": "Bolder", "wrap": True, "color": "Attention" if urgency >= 70 else "Warning" if urgency >= 50 else "Default"})
                    email_container["items"].append({"type": "FactSet", "facts": [
                        {"title": "From:", "value": e['sender_name']},
                        {"title": "🎯 Urgency:", "value": f"{urgency}/100"}
                    ]})
                    email_container["items"].append({"type": "TextBlock", "text": f"**Summary:** {e['smart_summary'][:150]}", "wrap": True})

                    if urgency >= 70:
                        smart_reply = generate_smart_reply(e, weather)
                        if smart_reply:
                            email_container["items"].append({"type": "TextBlock", "text": "🤖 **Smart Reply (Urgent):**", "weight": "Bolder", "separator": True, "spacing": "Small"})
                            email_container["items"].append({"type": "TextBlock", "text": smart_reply[:300], "wrap": True, "size": "Small"})

                    email_container["items"].append({"type": "ActionSet", "actions": [{"type": "Action.OpenUrl", "title": "📧 Open", "url": f"https://outlook.office.com/mail/inbox/id/{e['entry_id']}"}]})
                    card_body.append(email_container)

        if vessels_info:
            card_body.append({"type": "TextBlock", "text": f"🚢 Vessels Mentioned ({len(vessels_info)})", "size": "Large", "weight": "Bolder", "separator": True, "spacing": "Large"})
            for vname, vdata in list(vessels_info.items())[:3]:
                icon = "⛵" if vdata.get('identifier_type') == 'ENI' else "🚢"
                card_body.append({"type": "Container", "items": [
                    {"type": "TextBlock", "text": f"{icon} {vname}", "weight": "Bolder"},
                    {"type": "ActionSet", "actions": [{"type": "Action.OpenUrl", "title": "🔍 Track", "url": vdata['vessel_url']}]}
                ], "separator": True})

        card_body.append({"type": "TextBlock", "text": "📅 Today's Calendar", "size": "Large", "weight": "Bolder", "separator": True, "spacing": "Large"})
        if events:
            card_body.append({"type": "TextBlock", "text": f"**{len(events)} event(s) scheduled**", "weight": "Bolder"})
            for ev in events[:5]:
                card_body.append({"type": "Container", "items": [
                    {"type": "TextBlock", "text": f"🗓️ **{ev['subject']}**", "weight": "Bolder", "wrap": True},
                    {"type": "FactSet", "facts": [
                        {"title": "⏰ Time:", "value": f"{ev['start_time']} - {ev['end_time']}"},
                        {"title": "📍 Location:", "value": ev['location']}
                    ]}
                ], "separator": True})
        else:
            card_body.append({"type": "TextBlock", "text": "✅ No meetings scheduled", "color": "Good"})

        card_body.append({"type": "TextBlock", "text": f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | v20.0 FINAL", "size": "Small", "isSubtle": True, "separator": True})

        payload = {
            "type": "message",
            "attachments": [{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": card_body
                }
            }]
        }

        payload_str = json.dumps(payload)
        payload_size = len(payload_str.encode('utf-8'))
        log(f"📦 Payload size: {payload_size:,} bytes ({payload_size/1024:.1f} KB)")
        
        if payload_size > 28000:
            log("⚠️ WARNING: Payload exceeds 28KB - Teams may reject it!")
        
        log("📤 Sending to Teams...")
        try:
            response = requests.post(TEAMS_WEBHOOK_URL, json=payload, proxies=PROXIES, timeout=10, verify=False)
            response.raise_for_status()
            
            log(f"✅ Teams HTTP {response.status_code}: {response.text[:100]}")
            
            if response.status_code == 200:
                log("✅ Teams summary sent successfully!")
            else:
                log(f"⚠️ Teams returned {response.status_code}")
                
        except Exception as send_error:
            log(f"❌ Teams send error: {send_error}")
            if 'response' in locals():
                log(f"Response: {response.text}")

    except Exception as e:
        log(f"❌ Teams error: {e}")
        log_exception(e)

# =========================================================
# MAIN AGENT
# =========================================================

def run_summary_agent():
    log("=" * 60)
    log("🚀 Email Assistant v20.0 FINAL - INTELLIGENT VERSION")
    log("✅ Email-to-Checklist | ✅ Anchored Date | ✅ Status Display")
    log("=" * 60)


    try:
        weather = get_weather_conditions()
        emails = fetch_emails() or []

        pilot_status = load_pilot_status()
        vessels_info = collect_vessel_info(emails) if emails else {}

        timeline = load_timeline()
        log(f"📊 Timeline loaded: {len(timeline.get('vessels', []))} vessels")

        conflicts = detect_conflicts(timeline)
        log(f"⚠️ Detected {len(conflicts)} conflicts")

        # Excel-based updates
        checklists = update_checklists(timeline)
        checklists = cleanup_old_checklists(checklists, timeline)
        
        # NEW: Email-based updates
        checklists, delays_detected = update_checklists_from_emails(checklists, emails)
        
        checklist_summary = get_checklist_summary(checklists)
        log(f"📋 Active checklists: {checklist_summary['total']}")

        events = fetch_calendar() or []

        send_summary_to_teams(emails, events, weather, vessels_info, pilot_status, timeline, conflicts, checklist_summary, checklists, delays_detected)

        log("=" * 60)
        log(f"✅ Complete: {len(emails)} emails, {len(timeline.get('vessels', []))} vessels, {len(events)} events")
        if checklist_summary['total'] > 0:
            log(f"📋 {checklist_summary['total']} active checklists")
        if delays_detected:
            log(f"🚨 {len(delays_detected)} delay warnings detected")
        log("=" * 60)

    except Exception as e:
        log(f"Fatal error: {e}")
        log_exception(e)

def is_work_hours():
    return WORK_HOURS_START <= datetime.now().hour < WORK_HOURS_END

def scheduled_run():
    if is_work_hours():
        log("🕐 Scheduled run...")
        run_summary_agent()

if __name__ == "__main__":
    try:
        log("🤖 Email Assistant Starting - v20.0 FINAL")
        log("📧 Smart Email Parsing Enabled")
        log("📊 Anchored Date Logic Active")
        log("🎯 Status Display Integrated")
        run_summary_agent()
        schedule.every(RUN_INTERVAL_HOURS).hours.do(scheduled_run)
        log("✅ Agent running continuously (Ctrl+C to stop)")
        while True:
            schedule.run_pending()
            time.sleep(60)
    except KeyboardInterrupt:
        log("⏹️ Stopped by user")
    except Exception as e:
        log(f"CRITICAL: {e}")
        log_exception(e)