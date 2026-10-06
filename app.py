import os
import csv
import re
import shutil
import subprocess
import tempfile
import threading
import queue
import datetime
import platform
import time
import json
import traceback
import tkinter as tk
import pandas as pd
from tkinter import ttk
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import paramiko
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURATION
# ============================================================
APP_TITLE = "INFOBLOX DIG AUTOMATION By Suraj Joshi"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COMMANDS_FILE = os.path.join(
    BASE_DIR,
    "commands.txt"
)
OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "output"
)
SERVER_PROFILES_FILE = os.path.join(
    BASE_DIR,
    "server_profiles.json"
)


# ============================================================
# QUARTERLY APPLICATION LICENSE
# ============================================================
# The application contains a code for each quarter.
# The user enters the current quarter's code once; the valid
# quarter/code is stored locally and reused until the quarter ends.
QUARTERLY_LICENSE_CODES = {
    "2026-Q4": "IBX-9B4Y-3FZK-TRPN",
    "2027-Q1": "IBX-MVVY-C8QJ-YY8L",
    "2027-Q2": "IBX-DTEF-YR2Z-AETM",
    "2027-Q3": "IBX-TR2H-2XVA-AL8L",
    "2027-Q4": "IBX-7HXC-MPW2-N979",
    "2028-Q1": "IBX-482L-G9GW-U6AM",
    "2028-Q2": "IBX-K2XR-YGTL-W8SC",
    "2028-Q3": "IBX-G5KH-W2N9-ESHD",
    "2028-Q4": "IBX-4QA8-HNSU-W6E3",
    "2029-Q1": "IBX-7D7J-5R8D-UHD9",
    "2029-Q2": "IBX-Q9GB-FBCK-8FVT",
    "2029-Q3": "IBX-N6PN-UHNS-8VHB",
}

LICENSE_STORAGE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "Infoblox_DIG_Automation"
)
LICENSE_STORAGE_FILE = os.path.join(LICENSE_STORAGE_DIR, "license.json")


def get_current_license_period(today=None):
    """Return (quarter_key, quarter_end_date) for the current date."""
    today = today or datetime.date.today()
    year = today.year
    month = today.month

    if month <= 3:
        quarter = 1
        expiry = datetime.date(year, 3, 31)
    elif month <= 6:
        quarter = 2
        expiry = datetime.date(year, 6, 30)
    elif month <= 9:
        quarter = 3
        expiry = datetime.date(year, 9, 30)
    else:
        quarter = 4
        expiry = datetime.date(year, 12, 31)

    return f"{year}-Q{quarter}", expiry


def load_local_license():
    try:
        with open(LICENSE_STORAGE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return {}


def save_local_license(code, quarter_key):
    os.makedirs(LICENSE_STORAGE_DIR, exist_ok=True)
    data = {
        "quarter": quarter_key,
        "code": code,
        "activated_on": datetime.date.today().isoformat(),
    }
    temp_file = LICENSE_STORAGE_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(temp_file, LICENSE_STORAGE_FILE)


def ensure_quarterly_license(parent=None):
    """Validate the current quarter license before showing the main GUI."""
    quarter_key, expiry_date = get_current_license_period()
    expected_code = QUARTERLY_LICENSE_CODES.get(quarter_key)

    # Create a small temporary Tk root for the license dialogs. This is
    # deliberately separate from the main CustomTkinter application root
    # so the license prompt is always visible before the main GUI starts.
    dialog_root = None
    if parent is None:
        dialog_root = tk.Tk()
        dialog_root.withdraw()
        parent = dialog_root

    try:
        if not expected_code:
            messagebox.showerror(
                "License Not Configured",
                (
                    f"No application license code is configured for {quarter_key}.\n\n"
                    "Please contact the application administrator."
                ),
                parent=parent,
            )
            return False

        saved = load_local_license()
        saved_code = str(saved.get("code", "")).strip().upper()
        saved_quarter = str(saved.get("quarter", "")).strip()

        # A matching saved license is valid only for the current quarter.
        if saved_quarter == quarter_key and saved_code == expected_code:
            return True

        expiry_text = expiry_date.strftime("%d-%b-%Y")
        prompt = (
            f"Quarterly license required for {quarter_key}.\n\n"
            f"Valid until: {expiry_text}\n\n"
            "Enter the application code provided by the administrator:"
        )

        while True:
            code = simpledialog.askstring(
                "Infoblox DIG Automation License",
                prompt,
                parent=parent,
            )

            if code is None:
                return False

            code = code.strip().upper()
            if code == expected_code:
                try:
                    save_local_license(code, quarter_key)
                except OSError as exc:
                    messagebox.showerror(
                        "License Save Error",
                        f"The license was valid, but it could not be saved locally.\n\n{exc}",
                        parent=parent,
                    )
                    return False
                return True

            messagebox.showerror(
                "Invalid License Code",
                (
                    "The code entered is not valid for the current quarter.\n\n"
                    "Please enter the current quarterly code."
                ),
                parent=parent,
            )
    finally:
        if dialog_root is not None:
            try:
                dialog_root.destroy()
            except Exception:
                pass
os.makedirs(
    OUTPUT_DIR,
    exist_ok=True
)
# Active SSH channel used by the manual Infoblox pager button.
ACTIVE_SSH_CHANNEL = None
ACTIVE_SSH_CHANNEL_LOCK = threading.Lock()
# ============================================================
# MODERN UI THEME CONSTANTS
# ============================================================
# Color palette (dark modern theme)
COLOR_BG            = "#0f1419"   # app background
COLOR_SURFACE       = "#1a2029"   # card surface
COLOR_SURFACE_ALT   = "#232b36"   # elevated surface
COLOR_BORDER        = "#2d3743"   # subtle border
COLOR_ACCENT        = "#3b82f6"   # primary blue
COLOR_ACCENT_HOVER  = "#2563eb"
COLOR_SUCCESS       = "#10b981"   # green
COLOR_WARNING       = "#f59e0b"   # amber
COLOR_DANGER        = "#ef4444"   # red
COLOR_TEXT          = "#e6edf3"   # primary text
COLOR_TEXT_MUTED    = "#8b949e"   # secondary text
COLOR_PURPLE        = "#8b5cf6"
COLOR_CYAN          = "#06b6d4"
# Fonts
FONT_FAMILY = "Segoe UI"
def font(size=12, weight="normal"):
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)
# ============================================================
# UTILITY
# ============================================================
def now_string():
    return datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )
def clean_terminal_text(text):
    """
    Remove terminal control characters.
    Important for Infoblox output because SSH/TTY can
    produce backspace characters such as:
        +sh\x08ort
    which visually appears as:
        +short
    """
    if not text:
        return ""
    # Remove OSC terminal integration sequences first.
    # Ubuntu shells can emit OSC 3008 command/shell markers; if the
    # leading ESC is removed by the generic ANSI cleanup, the payload
    # (e.g. 3008;start=...) would otherwise remain visible.
    text = re.sub(
        r"\x1B\][^\x07\x1B]*(?:\x07|\x1B\\)",
        "",
        text
    )
    # Remove ANSI escape sequences
    text = re.sub(
        r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])",
        "",
        text
    )
    # Normalize terminal carriage returns so output is readable
    # instead of containing visible CR formatting artifacts.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Handle backspace sequences
    while "\x08" in text:
        text = re.sub(
            r".\x08",
            "",
            text
        )
    # Remove remaining control chars except CR/LF/TAB
    text = "".join(
        ch
        for ch in text
        if ch in "\r\n\t" or ord(ch) >= 32
    )
    return text
def normalize_command(command):
    """
    Normalize a DIG command for matching between
    Step 1 and Step 3.
    """
    if not command:
        return ""
    command = clean_terminal_text(command)
    command = command.strip()
    # Remove prompt if present
    if ">" in command:
        command = command.split(">")[-1].strip()
    # Remove +short
    command = re.sub(
        r"\s+\+short\b",
        "",
        command,
        flags=re.IGNORECASE
    )
    # Remove multiple spaces
    command = re.sub(
        r"\s+",
        " ",
        command
    )
    return command.strip()
def command_key(command):
    return normalize_command(command).lower()
# ============================================================
# PUTTY / PPK SUPPORT
# ============================================================
def find_puttygen():
    """Locate puttygen.exe."""
    candidates = [
        shutil.which("puttygen"),
        r"C:\Program Files\PuTTY\puttygen.exe",
        r"C:\Program Files (x86)\PuTTY\puttygen.exe",
        os.path.expandvars(
            r"%LOCALAPPDATA%\Programs\PuTTY\puttygen.exe"
        ),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None
def convert_ppk_to_openssh(
    ppk_file,
    passphrase=None
):
    """
    Convert PPK to temporary OpenSSH private key.
    Original PPK is never modified.
    """
    puttygen = find_puttygen()
    if not puttygen:
        raise Exception(
            "puttygen.exe was not found.\n\n"
            "Please install PuTTY or add puttygen.exe "
            "to your Windows PATH."
        )
    temp_dir = tempfile.mkdtemp(
        prefix="infoblox_key_"
    )
    openssh_key = os.path.join(
        temp_dir,
        "converted_key"
    )
    command = [
        puttygen,
        ppk_file,
        "-O",
        "private-openssh",
        "-o",
        openssh_key
    ]
    if passphrase:
        command.extend([
            "--old-passphrase",
            passphrase
        ])
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        creationflags=(
            subprocess.CREATE_NO_WINDOW
            if platform.system() == "Windows"
            else 0
        )
    )
    if result.returncode != 0:
        error = (
            result.stderr.strip()
            or result.stdout.strip()
            or "Unknown PuTTYgen error."
        )
        shutil.rmtree(
            temp_dir,
            ignore_errors=True
        )
        raise Exception(
            "PPK conversion failed:\n\n"
            + error
        )
    if not os.path.exists(openssh_key):
        shutil.rmtree(
            temp_dir,
            ignore_errors=True
        )
        raise Exception(
            "PuTTYgen completed but the converted "
            "private key was not created."
        )
    return openssh_key, temp_dir
# ============================================================
# PRIVATE KEY LOADING
# ============================================================
def load_private_key(
    key_file,
    passphrase=None
):
    """
    Load either:
        .pem
        .ppk
    Returns:
        (paramiko_key, temporary_directory)
    temporary_directory is None for PEM.
    """
    if not key_file:
        raise Exception(
            "Please select a private key."
        )
    key_file = os.path.abspath(
        key_file
    )
    if not os.path.isfile(key_file):
        raise Exception(
            f"Private key not found:\n{key_file}"
        )
    extension = os.path.splitext(
        key_file
    )[1].lower()
    temporary_directory = None
    try:
        # ----------------------------------------------------
        # PPK
        # ----------------------------------------------------
        if extension == ".ppk":
            key_file, temporary_directory = (
                convert_ppk_to_openssh(
                    key_file,
                    passphrase
                )
            )
        # ----------------------------------------------------
        # PEM / OpenSSH
        # ----------------------------------------------------
        key_classes = []
        if hasattr(paramiko, "RSAKey"):
            key_classes.append(
                paramiko.RSAKey
            )
        if hasattr(paramiko, "ECDSAKey"):
            key_classes.append(
                paramiko.ECDSAKey
            )
        if hasattr(paramiko, "Ed25519Key"):
            key_classes.append(
                paramiko.Ed25519Key
            )
        if hasattr(paramiko, "DSSKey"):
            key_classes.append(
                paramiko.DSSKey
            )
        last_error = None
        for key_class in key_classes:
            try:
                key = key_class.from_private_key_file(
                    key_file,
                    password=passphrase
                )
                return (
                    key,
                    temporary_directory
                )
            except Exception as error:
                last_error = error
        raise Exception(
            "Unable to load the private key.\n\n"
            f"File:\n{key_file}\n\n"
            f"Last error:\n{last_error}"
        )
    except Exception:
        if temporary_directory:
            shutil.rmtree(
                temporary_directory,
                ignore_errors=True
            )
        raise
# ============================================================
# READ COMMANDS
# ============================================================
def read_commands(commands_file=None, trailing_option=""):
    if not commands_file:
        commands_file = COMMANDS_FILE
    # A non-empty GUI trailing option is added to generated DIG commands.
    # When the field is empty, no trailing option is added.
    trailing_option = (trailing_option or "").strip()
    if not os.path.exists(
        commands_file
    ):
        raise Exception(
            f"Input file not found:\n\n"
            f"{commands_file}"
        )
    # ========================================================
    # EXCEL / CSV INPUT
    # ========================================================
    #
    # Expected columns:
    #   fqdn*
    #   forward_to
    #
    # Example:
    #   www.kyc.com
    #   gtm01.wallst.com/209.234.234.43,gtm02.wallst.com/209.234.230.6
    #
    # This creates:
    #   dig @209.234.234.43 www.kyc.com
    #   dig @209.234.230.6 www.kyc.com
    #
    # ========================================================
    extension = os.path.splitext(
        commands_file
    )[1].lower()
    if extension in (".xlsx", ".xlsm", ".csv"):
        commands = []
        if extension in (".xlsx", ".xlsm"):
            try:
                workbook = load_workbook(
                    commands_file,
                    read_only=True,
                    data_only=True
                )
            except Exception as error:
                raise Exception(
                    "Unable to read Excel file.\n\n"
                    f"{error}"
                )
            try:
                sheet = workbook.active
                rows = sheet.iter_rows(
                    values_only=True
                )
                try:
                    headers = next(rows)
                except StopIteration:
                    headers = []
                header_map = {}
                for index, header in enumerate(headers):
                    if header is None:
                        continue
                    header_name = str(
                        header
                    ).strip().lower()
                    header_map[header_name] = index
                if "fqdn*" not in header_map:
                    raise Exception(
                        "Excel file must contain an 'fqdn*' column."
                    )
                if "forward_to" not in header_map:
                    raise Exception(
                        "Excel file must contain a 'forward_to' column."
                    )
                fqdn_index = header_map["fqdn*"]
                forward_to_index = header_map["forward_to"]
                for row in rows:
                    if not row:
                        continue
                    fqdn = (
                        str(row[fqdn_index]).strip()
                        if fqdn_index < len(row)
                        and row[fqdn_index] is not None
                        else ""
                    )
                    forward_to = (
                        str(row[forward_to_index]).strip()
                        if forward_to_index < len(row)
                        and row[forward_to_index] is not None
                        else ""
                    )
                    if not fqdn or not forward_to:
                        continue
                    for target in forward_to.split(","):
                        target = target.strip()
                        if "/" not in target:
                            continue
                        ip = target.rsplit(
                            "/",
                            1
                        )[1].strip()
                        if not ip:
                            continue
                        command = f"dig @{ip} {fqdn}"
                        if trailing_option:
                            command += f" {trailing_option}"
                        commands.append(command)
            finally:
                workbook.close()
        else:
            # CSV input
            try:
                with open(
                    commands_file,
                    "r",
                    encoding="utf-8-sig",
                    newline=""
                ) as file:
                    reader = csv.reader(file)
                    try:
                        headers = next(reader)
                    except StopIteration:
                        headers = []
                    header_map = {}
                    for index, header in enumerate(headers):
                        header_name = str(
                            header
                        ).strip().lower()
                        header_map[header_name] = index
                    if "fqdn*" not in header_map:
                        raise Exception(
                            "CSV file must contain an 'fqdn*' column."
                        )
                    if "forward_to" not in header_map:
                        raise Exception(
                            "CSV file must contain a 'forward_to' column."
                        )
                    fqdn_index = header_map["fqdn*"]
                    forward_to_index = header_map["forward_to"]
                    for row in reader:
                        if not row:
                            continue
                        fqdn = (
                            row[fqdn_index].strip()
                            if fqdn_index < len(row)
                            else ""
                        )
                        forward_to = (
                            row[forward_to_index].strip()
                            if forward_to_index < len(row)
                            else ""
                        )
                        if not fqdn or not forward_to:
                            continue
                        for target in forward_to.split(","):
                            target = target.strip()
                            if "/" not in target:
                                continue
                            ip = target.rsplit(
                                "/",
                                1
                            )[1].strip()
                            if not ip:
                                continue
                            command = f"dig @{ip} {fqdn}"
                            if trailing_option:
                                command += f" {trailing_option}"
                            commands.append(command)
            except UnicodeDecodeError:
                raise Exception(
                    "Unable to read CSV file as UTF-8."
                )
        return commands
    # ========================================================
    # TXT INPUT - EXISTING LOGIC
    # ========================================================
    commands = []
    with open(
        commands_file,
        "r",
        encoding="utf-8"
    ) as file:
        for line in file:
            line = line.strip()
            if not line:
                continue
            # ------------------------------------------------
            # Allow both formats in the command file:
            #
            #   dig @10.2.3.10 ihs.com
            #   dig ttngl.com
            #   tatate.com
            #
            # If the line is only a domain, automatically
            # convert it to a DIG command.
            # ------------------------------------------------
            if line.lower().startswith(
                "dig "
            ):
                command = line
            else:
                command = "dig " + line
            # Apply the GUI-selected trailing option only when provided.
            command_parts = command.split()
            if (
                trailing_option
                and (
                    len(command_parts) == 2
                    or (
                        len(command_parts) == 3
                        and command_parts[1].startswith("@")
                    )
                )
            ):
                command += f" {trailing_option}"
            commands.append(command)
    return commands
# ============================================================
# SSH EXECUTION
# ============================================================
def execute_ssh(
    server,
    username,
    port,
    authentication,
    password,
    key_file,
    key_passphrase,
    commands,
    log_callback,
    keep_connection=False,
    progress_callback=None
):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(
        paramiko.AutoAddPolicy()
    )
    private_key = None
    temp_key_dir = None
    keep_open = False
    try:
        log_callback(
            f"Connecting to {server}:{port}..."
        )
        # ----------------------------------------------------
        # PASSWORD AUTH
        # ----------------------------------------------------
        if authentication == "Password":
            client.connect(
                hostname=server,
                port=int(port),
                username=username,
                password=password,
                timeout=20,
                banner_timeout=20,
                auth_timeout=20,
                look_for_keys=False,
                allow_agent=False
            )
        # ----------------------------------------------------
        # KEY AUTH
        # ----------------------------------------------------
        else:
            private_key, temp_key_dir = (
                load_private_key(
                    key_file,
                    key_passphrase
                )
            )
            log_callback(
                "Private key loaded successfully."
            )
            client.connect(
                hostname=server,
                port=int(port),
                username=username,
                pkey=private_key,
                timeout=20,
                banner_timeout=20,
                auth_timeout=20,
                look_for_keys=False,
                allow_agent=False
            )
        log_callback(
            f"Connected to {server}"
        )
        # ----------------------------------------------------
        # OPEN SHELL
        # ----------------------------------------------------
        channel = client.invoke_shell(
            term="xterm",
            width=160,
            height=200
        )
        global ACTIVE_SSH_CHANNEL
        with ACTIVE_SSH_CHANNEL_LOCK:
            ACTIVE_SSH_CHANNEL = channel
        channel.settimeout(
            1.0
        )
        # Give Infoblox CLI time to initialize
        time.sleep(2)
        # Clear initial output
        initial = ""
        while True:
            try:
                if channel.recv_ready():
                    data = channel.recv(
                        65535
                    ).decode(
                        "utf-8",
                        errors="replace"
                    )
                    initial += data
                else:
                    break
            except Exception:
                break
        # ----------------------------------------------------
        # RUN COMMANDS ONE AT A TIME
        # ----------------------------------------------------
        # Important: Infoblox can open a pager for a large DIG
        # result. Sending all commands in one batch causes the
        # remaining commands to stop advancing when the pager is
        # cancelled with q. We therefore send the next DIG only
        # after the Infoblox prompt returns.
        log_callback(
            f"Executing {len(commands)} DIG commands..."
        )
        output = initial
        command_index = 0
        exit_sent = False
        pager_handled = False
        start_time = time.time()
        last_data_time = time.time()
        max_runtime = max(120, len(commands) * 15)
        if commands:
            channel.send(
                commands[0] + "\n"
            )
            command_index = 1
        while True:
            if channel.recv_ready():
                data = channel.recv(
                    65535
                ).decode(
                    "utf-8",
                    errors="replace"
                )
                output += data
                last_data_time = time.time()
                log_callback(
                    data,
                    raw=True
                )
                # Automatically cancel the Infoblox pager when it appears.
                # This is checked for every DIG command. The flag resets
                # when the next DIG command is sent.
                if (
                    not pager_handled
                    and re.search(
                        r"Enter <return> for next page or q<return> to cancel the command\.?",
                        clean_terminal_text(data),
                        flags=re.IGNORECASE
                    )
                ):
                    channel.send("q\n")
                    pager_handled = True
                    last_data_time = time.time()
                    log_callback(
                        "[AUTO] Infoblox pager detected - sent q + Enter. Continuing..."
                    )
            else:
                time.sleep(0.1)
            elapsed = (
                time.time()
                - start_time
            )
            cleaned = clean_terminal_text(output)
            prompt_returned = re.search(
                r"(?:^|\n)[^\n]*>\s*$",
                cleaned,
                flags=re.MULTILINE
            )
            if (
                prompt_returned
                and time.time() - last_data_time > 0.5
            ):
                if command_index < len(commands):
                    # The previous DIG finished. Report progress before
                    # starting the next DIG.
                    if progress_callback:
                        progress_callback(
                            command_index,
                            len(commands)
                        )
                    # Send exactly the next command. This also handles
                    # manual Q: q returns Infoblox to this prompt.
                    channel.send(
                        commands[command_index] + "\n"
                    )
                    command_index += 1
                    pager_handled = False
                    last_data_time = time.time()
                    continue
                # The final DIG has completed.
                if progress_callback:
                    progress_callback(
                        len(commands),
                        len(commands)
                    )
                if keep_connection:
                    break
                if not exit_sent:
                    channel.send("exit\n")
                    exit_sent = True
                    last_data_time = time.time()
                    continue
            if (
                exit_sent
                and re.search(
                    r"(^|\n).*\>\s*exit\s*",
                    output
                )
                and time.time() - last_data_time > 1.5
            ):
                break
            if elapsed > max_runtime:
                log_callback(
                    "SSH execution timeout reached."
                )
                break
        cleaned_output = clean_terminal_text(
            output
        )
        if keep_connection:
            keep_open = True
            return (
                cleaned_output,
                client,
                channel,
                temp_key_dir
            )
        return cleaned_output
    except Exception as exc:
        log_callback(
            f"SSH ERROR: {exc}"
        )
        return ""
    finally:
        if not keep_open:
            with ACTIVE_SSH_CHANNEL_LOCK:
                if 'channel' in locals() and ACTIVE_SSH_CHANNEL is channel:
                    ACTIVE_SSH_CHANNEL = None
            try:
                client.close()
            except Exception:
                pass
            if temp_key_dir:
                shutil.rmtree(
                    temp_key_dir,
                    ignore_errors=True
                )
def execute_on_existing_session(
    channel,
    commands,
    log_callback,
    progress_callback=None
):
    """Run commands one at a time on the already authenticated SSH session."""
    if not commands:
        return ""
    global ACTIVE_SSH_CHANNEL
    with ACTIVE_SSH_CHANNEL_LOCK:
        ACTIVE_SSH_CHANNEL = channel
    log_callback(
        f"Executing {len(commands)} DIG commands on existing SSH session..."
    )
    output = ""
    command_index = 0
    pager_handled = False
    start_time = time.time()
    last_data_time = time.time()
    max_runtime = max(120, len(commands) * 15)
    # Send only the first command. The next command is sent only after
    # the Infoblox prompt returns. This is what allows the manual Q
    # button to cancel a pager and then continue with the next DIG.
    channel.send(
        commands[0] + "\n"
    )
    command_index = 1
    while True:
        if channel.recv_ready():
            data = channel.recv(
                65535
            ).decode(
                "utf-8",
                errors="replace"
            )
            output += data
            last_data_time = time.time()
            log_callback(
                data,
                raw=True
            )
            # Automatically cancel the Infoblox pager when it appears.
            # This is checked for every DIG command. The flag resets
            # when the next DIG command is sent.
            # Detect the pager only in the data received for the
            # current DIG command. Do NOT search the complete output,
            # otherwise a pager from an earlier command would trigger q
            # again on every later command.
            if (
                not pager_handled
                and re.search(
                    r"Enter <return> for next page or q<return> to cancel the command\.?",
                    clean_terminal_text(data),
                    flags=re.IGNORECASE
                )
            ):
                channel.send("q\n")
                pager_handled = True
                last_data_time = time.time()
                log_callback(
                    "[AUTO] Infoblox pager detected - sent q + Enter. Continuing..."
                )
        else:
            time.sleep(0.1)
        elapsed = time.time() - start_time
        cleaned = clean_terminal_text(output)
        prompt_returned = re.search(
            r"(?:^|\n)[^\n]*>\s*$",
            cleaned,
            flags=re.MULTILINE
        )
        if (
            prompt_returned
            and time.time() - last_data_time > 0.5
        ):
            if command_index < len(commands):
                if progress_callback:
                    progress_callback(
                        command_index,
                        len(commands)
                    )
                channel.send(
                    commands[command_index] + "\n"
                )
                command_index += 1
                pager_handled = False
                last_data_time = time.time()
                continue
            if progress_callback:
                progress_callback(
                    len(commands),
                    len(commands)
                )
            break
        if elapsed > max_runtime:
            log_callback(
                "SSH execution timeout reached."
            )
            break
    return clean_terminal_text(output)
# ============================================================
# PARSE DIG BLOCKS
# ============================================================
def parse_dig_blocks(text):
    """Parse only real Infoblox prompt commands; ignore DiG banner text."""
    text = clean_terminal_text(text or "")
    blocks = []
    current_command = None
    current_output = []
    for line in text.splitlines():
        clean_line = clean_terminal_text(line).strip()
        # Only an Infoblox prompt followed by the word DIG starts a block.
        match = re.search(r"(?:^|\s)>\s*(dig\b.*)$", clean_line, re.IGNORECASE)
        if match:
            if current_command:
                blocks.append((current_command, "\n".join(current_output)))
            current_command = match.group(1).strip()
            current_output = []
            continue
        if re.search(r"(?:^|\s)>\s*exit\s*$", clean_line, re.IGNORECASE):
            if current_command:
                blocks.append((current_command, "\n".join(current_output)))
            current_command = None
            current_output = []
            continue
        if current_command:
            current_output.append(line)
    if current_command:
        blocks.append((current_command, "\n".join(current_output)))
    return blocks
def detect_actual_dns_status(output):
    """Return the real DNS status from a dig response."""
    output = clean_terminal_text(output or "")
    stripped = output.strip()
    lower = stripped.lower()
    if any(x in lower for x in (
        "connection timed out",
        "no servers could be reached",
        "communications error",
        "timed out",
    )):
        return "CONNECTION TIMED OUT"
    status_match = re.search(
        r"\bstatus\s*:\s*([A-Z]+)",
        stripped,
        flags=re.IGNORECASE
    )
    if status_match:
        return status_match.group(1).upper()
    ipv4_pattern = (
        r"\b(?:25[0-5]|2[0-4]\d|1?\d?\d)"
        r"(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}\b"
    )
    if re.search(ipv4_pattern, stripped):
        return "NOERROR"
    if re.search(r"\b[0-9a-fA-F]*:[0-9a-fA-F:]+\b", stripped):
        return "NOERROR"
    return "UNSUCCESSFUL" if not stripped else "OTHER"
def parse_step3_statuses(step3_output, unsuccessful_commands):
    """Extract actual Step 3 statuses in retry-command order.
    This ignores '+noedns ...' and all other DiG banner text.
    """
    text = clean_terminal_text(step3_output or "")
    statuses = []
    for line in text.splitlines():
        lower = line.lower()
        if "connection timed out" in lower or "no servers could be reached" in lower:
            statuses.append("CONNECTION TIMED OUT")
            continue
        match = re.search(r"\bstatus\s*:\s*([A-Z]+)", line, re.IGNORECASE)
        if match:
            statuses.append(match.group(1).upper())
    return statuses[:len(unsuccessful_commands)]
def parse_actual_results(step_output):
    result = {}
    for command, output in parse_dig_blocks(step_output):
        key = command_key(command)
        result[key] = {
            "command": normalize_command(command),
            "status": detect_actual_dns_status(output),
            "output": output
        }
    return result
def extract_unsuccessful(
    step1_output
):
    blocks = parse_dig_blocks(
        step1_output
    )
    unsuccessful = []
    for command, output in blocks:
        status = detect_actual_dns_status(
            output
        )
        # Only truly unsuccessful commands are retried in Step 3.
        # A connection timeout is already a final result.
        if status == "UNSUCCESSFUL":
            unsuccessful.append(
                normalize_command(
                    command
                )
            )
    return unsuccessful
# ============================================================
# STEP 3 EXECUTION
# ============================================================
def run_step3(
    server,
    username,
    port,
    authentication,
    password,
    key_file,
    key_passphrase,
    unsuccessful_commands,
    log_callback,
    session=None,
    progress_callback=None
):
    if not unsuccessful_commands:
        return ""
    commands = [
        cmd
        for cmd in unsuccessful_commands
        if cmd.strip()
    ]
    # Reuse the Step 1 authenticated SSH session when available.
    if session:
        return execute_on_existing_session(
            channel=session[1],
            commands=commands,
            log_callback=log_callback,
            progress_callback=progress_callback
        )
    # Backward-compatible fallback: create a new SSH session.
    return execute_ssh(
        server=server,
        username=username,
        port=port,
        authentication=authentication,
        password=password,
        key_file=key_file,
        key_passphrase=key_passphrase,
        commands=commands,
        log_callback=log_callback,
        progress_callback=progress_callback
    )
# ============================================================
# FINAL TXT
# ============================================================
def create_final_file(
    server,
    username,
    authentication,
    commands,
    step1_output,
    unsuccessful,
    step3_output
):
    timestamp = datetime.datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )
    final_file = os.path.join(
        OUTPUT_DIR,
        f"infoblox_dig_final_{timestamp}.txt"
    )
    with open(
        final_file,
        "w",
        encoding="utf-8"
    ) as file:
        file.write(
            "=" * 70
            + "\n"
        )
        file.write(
            "INFOBLOX DIG AUTOMATION RESULT\n"
        )
        file.write(
            "=" * 70
            + "\n\n"
        )
        file.write(
            f"Server   : {server}\n"
        )
        file.write(
            f"Username : {username}\n"
        )
        file.write(
            f"Generated: {now_string()}\n"
        )
        file.write(
            f"Auth     : {authentication}\n"
        )
        file.write(
            "\n"
        )
        # ----------------------------------------------------
        # STEP 1
        # ----------------------------------------------------
        file.write(
            "=" * 70
            + "\n"
        )
        file.write(
            "STEP 1 - ALL DIG RESULTS\n"
        )
        file.write(
            "=" * 70
            + "\n\n"
        )
        file.write(
            step1_output
        )
        file.write(
            "\n\n"
        )
        # ----------------------------------------------------
        # STEP 2
        # ----------------------------------------------------
        file.write(
            "=" * 70
            + "\n"
        )
        file.write(
            "STEP 2 - UNSUCCESSFUL COMMANDS\n"
        )
        file.write(
            "=" * 70
            + "\n\n"
        )
        for command in unsuccessful:
            file.write(
                command + "\n"
            )
        # ----------------------------------------------------
        # STEP 3
        # ----------------------------------------------------
        file.write(
            "\n"
            + "=" * 70
            + "\n"
        )
        file.write(
            "STEP 3 - RETRY RESULTS\n"
        )
        file.write(
            "=" * 70
            + "\n\n"
        )
        if step3_output:
            file.write(
                step3_output
            )
        else:
            file.write(
                "No commands required retry.\n"
            )
        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------
        file.write(
            "\n\n"
            + "=" * 70
            + "\n"
        )
        file.write(
            "SUMMARY\n"
        )
        file.write(
            "=" * 70
            + "\n"
        )
        file.write(
            f"Total commands       : {len(commands)}\n"
        )
        file.write(
            f"Unsuccessful commands: {len(unsuccessful)}\n"
        )
        file.write(
            "Step 1              : Completed\n"
        )
        file.write(
            "Step 2              : Completed\n"
        )
        file.write(
            "Step 3              : Completed\n"
        )
    return final_file
# ============================================================
# EXCEL
# ============================================================
def parse_actual_results(
    step_output
):
    result = {}
    blocks = parse_dig_blocks(
        step_output
    )
    for command, output in blocks:
        key = command_key(
            command
        )
        status = detect_actual_dns_status(
            output
        )
        result[key] = {
            "command": normalize_command(
                command
            ),
            "status": status,
            "output": output
        }
    return result
def create_excel_result(commands, step1_output, step3_output, unsuccessful, server):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_file = os.path.join(
        OUTPUT_DIR,
        f"infoblox_dig_report_{timestamp}.xlsx"
    )
    step1_results = parse_actual_results(step1_output)
    step3_statuses = parse_step3_statuses(step3_output, unsuccessful)
    print("\n========== STEP 3 PARSED STATUS ==========")
    for i, command in enumerate(unsuccessful):
        status = step3_statuses[i] if i < len(step3_statuses) else "NO RESULT"
        print(f"{normalize_command(command)}  =>  {status}")
    print("==========================================\n")
    # Step 3 always wins for commands that were retried.
    step3_by_key = {}
    for i, command in enumerate(unsuccessful):
        if i < len(step3_statuses):
            step3_by_key[command_key(command)] = step3_statuses[i]
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["Server / IP", "Command", "Status"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    final_rows = []
    for original_command in commands:
        clean_command = normalize_command(original_command)
        key = command_key(clean_command)
        if key in step3_by_key:
            final_status = step3_by_key[key]
        else:
            step1 = step1_results.get(key)
            final_status = step1["status"] if step1 else "NO RESULT"
        final_rows.append((clean_command, final_status))
        ws.append([server, clean_command, final_status])
    summary = wb.create_sheet("Summary")
    summary["A1"] = "INFOBLOX DIG SUMMARY"
    summary["A1"].font = Font(bold=True, size=14)
    summary["A3"] = "Total Commands"
    summary["B3"] = len(commands)
    summary["A4"] = "Retried Commands"
    summary["B4"] = len(unsuccessful)
    summary["A8"] = "FINAL DNS STATUS"
    summary["A8"].font = Font(bold=True, size=13)
    summary["A10"] = "Status"
    summary["B10"] = "Count"
    status_counts = {}
    for _, status in final_rows:
        status_counts[status] = status_counts.get(status, 0) + 1
    for cell in summary[10]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    row = 11
    preferred_order = [
        "NOERROR", "NXDOMAIN", "SERVFAIL", "REFUSED",
        "CONNECTION TIMED OUT", "UNSUCCESSFUL", "OTHER", "NO RESULT"
    ]
    added = set()
    for status in preferred_order:
        if status in status_counts:
            summary.cell(row=row, column=1, value=status)
            summary.cell(row=row, column=2, value=status_counts[status])
            added.add(status)
            row += 1
    for status, count in sorted(status_counts.items()):
        if status not in added:
            summary.cell(row=row, column=1, value=status)
            summary.cell(row=row, column=2, value=count)
            row += 1
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 28
    summary.column_dimensions["A"].width = 35
    summary.column_dimensions["B"].width = 15
    wb.save(excel_file)
    return excel_file
# ============================================================
# MODERN UI COMPONENTS
# ============================================================
class ModernCard(ctk.CTkFrame):
    """A rounded card with subtle border and optional title (grid-based)."""
    def __init__(self, master, title=None, **kwargs):
        super().__init__(
            master,
            fg_color=COLOR_SURFACE,
            corner_radius=16,
            border_width=1,
            border_color=COLOR_BORDER,
            **kwargs
        )
        self._title = title
        self.title_label = None
        # Reserve row 0 for the title (if any). Children should use row>=1.
        if title:
            self.title_label = ctk.CTkLabel(
                self,
                text=title,
                font=font(12, "bold"),
                text_color=COLOR_TEXT_MUTED,
                anchor="w"
            )
            self.title_label.grid(
                row=0,
                column=0,
                columnspan=10,
                padx=16,
                pady=(12, 2),
                sticky="w"
            )
class StatusPill(ctk.CTkFrame):
    """A status indicator with colored dot and label."""
    def __init__(self, master, title, initial_text="Waiting", color=COLOR_TEXT_MUTED):
        super().__init__(
            master,
            fg_color=COLOR_SURFACE_ALT,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER
        )
        self.color = color
        self.dot = ctk.CTkLabel(
            self,
            text="●",
            font=font(16),
            text_color=color
        )
        self.dot.pack(side="left", padx=(14, 6), pady=10)
        self.text_label = ctk.CTkLabel(
            self,
            text=initial_text,
            font=font(12, "bold"),
            text_color=COLOR_TEXT,
            anchor="w"
        )
        self.text_label.pack(side="left", padx=(0, 14), pady=10)
        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=font(10),
            text_color=COLOR_TEXT_MUTED,
            anchor="e"
        )
        self.title_label.pack(side="right", padx=(0, 14), pady=10)
    def set_status(self, text, color):
        self.dot.configure(text_color=color)
        self.text_label.configure(text=text)
        self.color = color
# ============================================================
# APPLICATION
# ============================================================
class InfobloxDigApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(fg_color=COLOR_BG)
        # Window sizing
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        initial_w = min(1280, max(900, screen_w - 80))
        initial_h = min(900, max(680, screen_h - 100))
        self.root.geometry(f"{initial_w}x{initial_h}")
        self.root.minsize(880, 620)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.log_queue = queue.Queue()
        self.last_txt = None
        self.last_excel = None
        self._last_responsive_width = 0
        self._last_quick_width = 0
        self._resize_job = None
        # Live output window
        self.console = None
        self.live_output_window = None
        self.live_output_buffer = ""
        # Interactive SSH Terminal
        self.terminal_window = None
        self.terminal_text = None
        self.terminal_command_entry = None
        self.terminal_status_label = None
        self.terminal_connect_button = None
        self.terminal_disconnect_button = None
        self.terminal_client = None
        self.terminal_channel = None
        self.terminal_reader_thread = None
        self.terminal_connected = False
        self.terminal_output_buffer = ""
        self.terminal_output_lock = threading.Lock()
        self.terminal_temp_key_dir = None
        # Status pill references (populated in build_gui)
        self._status_pills = {}
        # FQDN tab responsive state
        self._fqdn_last_width = 0
        self._fqdn_resize_job = None
        self.build_gui()
        self.process_log_queue()
    # ========================================================
    # GUI
    # ========================================================
    def build_gui(self):
        # ====================================================
        # ROOT GRID — only the content row expands
        # ====================================================
        self.root.grid_rowconfigure(0, weight=0)   # header: fixed
        self.root.grid_rowconfigure(1, weight=1)   # content: expands
        self.root.grid_columnconfigure(0, weight=1)
        # ====================================================
        # HEADER BAR
        # ====================================================
        header = ctk.CTkFrame(
            self.root,
            fg_color=COLOR_SURFACE,
            corner_radius=0,
            height=78,
            border_width=0
        )
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)
        # Left: brand mark
        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="w", padx=(22, 0), pady=14)
        logo = ctk.CTkLabel(
            brand,
            text="⬢",
            font=ctk.CTkFont(family=FONT_FAMILY, size=30, weight="bold"),
            text_color=COLOR_ACCENT
        )
        logo.pack(side="left", padx=(0, 12))
        title_box = ctk.CTkFrame(brand, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(
            title_box,
            text="INFOBLOX DIG AUTOMATION",
            font=font(19, "bold"),
            text_color=COLOR_TEXT,
            anchor="w"
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_box,
            text="SSH  •  DIG  •  Validation  •  Reports",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).pack(anchor="w")
        # Right: quick actions
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=2, sticky="e", padx=(0, 22), pady=14)
        self.live_output_button = ctk.CTkButton(
            actions,
            text="⬒  Live Output",
            width=140,
            height=36,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(12, "bold"),
            command=self.open_live_output
        )
        self.live_output_button.pack(side="right", padx=(8, 0))
        self.ssh_terminal_button = ctk.CTkButton(
            actions,
            text="⌨  SSH Terminal",
            width=140,
            height=36,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(12, "bold"),
            command=self.open_ssh_terminal
        )
        self.ssh_terminal_button.pack(side="right", padx=(8, 0))
        self.open_button = ctk.CTkButton(
            actions,
            text="📁  Output",
            width=120,
            height=36,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(12, "bold"),
            command=self.open_output
        )
        self.open_button.pack(side="right")
        # ====================================================
        # MAIN TABS
        # Existing DIG automation UI remains unchanged inside
        # the first tab. The second tab is the additive FQDN
        # Search + DIG feature.
        # ====================================================
        self.tabview = ctk.CTkTabview(
            self.root,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_ALT,
            segmented_button_unselected_hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            corner_radius=0
        )
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.tabview.add("DIG Automation")
        self.tabview.add("FQDN Search Tool")
        self.tabview._segmented_button.configure(
            font=font(12, "bold")
        )


        self.content = ctk.CTkScrollableFrame(
            self.tabview.tab("DIG Automation"),
            fg_color="transparent",
            corner_radius=0
        )
        self.content.pack(fill="both", expand=True)
        self.content.grid_columnconfigure(0, weight=1)


        self.build_fqdn_tab()
        # ====================================================
        # CONNECTION CARD
        # ====================================================
        conn_card = ModernCard(self.content, title="CONNECTION")
        conn_card.grid(row=0, column=0, sticky="ew", padx=22, pady=(14, 10))
        conn_card.grid_columnconfigure(1, weight=1)
        conn_card.grid_columnconfigure(3, weight=1)
        # --- Row: Server / Username ---
        ctk.CTkLabel(
            conn_card,
            text="Server / IP",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=1, column=0, padx=(16, 8), pady=(8, 4), sticky="w")
        self.server_entry = self._entry(conn_card)
        self.server_entry.grid(row=1, column=1, padx=(0, 14), pady=(8, 4), sticky="ew")
        self.server_entry.insert(0, "eudefra11dns02.info.corp")
        ctk.CTkLabel(
            conn_card,
            text="Username",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=1, column=2, padx=(0, 8), pady=(8, 4), sticky="w")
        self.username_entry = self._entry(conn_card)
        self.username_entry.grid(row=1, column=3, padx=(0, 16), pady=(8, 4), sticky="ew")
        self.username_entry.insert(0, "admin")
        # --- Row: Port / Auth ---
        ctk.CTkLabel(
            conn_card,
            text="SSH Port",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=2, column=0, padx=(16, 8), pady=4, sticky="w")
        self.port_entry = self._entry(conn_card, width=110)
        self.port_entry.grid(row=2, column=1, padx=(0, 14), pady=4, sticky="w")
        self.port_entry.insert(0, "22")
        ctk.CTkLabel(
            conn_card,
            text="Authentication",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=2, column=2, padx=(0, 8), pady=4, sticky="w")
        self.auth_var = ctk.StringVar(value="Password")
        self.auth_menu = ctk.CTkOptionMenu(
            conn_card,
            variable=self.auth_var,
            values=["Password", "SSH Key"],
            command=self.authentication_changed,
            width=180,
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            button_color=COLOR_ACCENT,
            button_hover_color=COLOR_ACCENT_HOVER,
            dropdown_fg_color=COLOR_SURFACE_ALT,
            dropdown_hover_color=COLOR_ACCENT,
            font=font(12)
        )
        self.auth_menu.grid(row=2, column=3, padx=(0, 16), pady=4, sticky="ew")
        # --- Row: Password ---
        self.password_label = ctk.CTkLabel(
            conn_card,
            text="SSH Password",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.password_label.grid(row=3, column=0, padx=(16, 8), pady=4, sticky="w")
        self.password_entry = self._entry(conn_card, show="*")
        self.password_entry.grid(row=3, column=1, columnspan=3, padx=(0, 16), pady=4, sticky="ew")
        # --- Key widgets (hidden initially) ---
        self.key_label = ctk.CTkLabel(
            conn_card,
            text="Private Key",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.key_entry = self._entry(conn_card)
        self.key_browse = ctk.CTkButton(
            conn_card,
            text="Browse",
            width=100,
            height=38,
            corner_radius=10,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            font=font(12, "bold"),
            command=self.browse_key
        )
        self.key_passphrase_label = ctk.CTkLabel(
            conn_card,
            text="Key Passphrase",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.key_passphrase_entry = self._entry(conn_card, show="*")
        # ====================================================
        # INPUT CARD
        # ====================================================
        input_card = ModernCard(self.content, title="INPUT")
        input_card.grid(row=1, column=0, sticky="ew", padx=22, pady=10)
        input_card.grid_columnconfigure(1, weight=1)
        # Command File
        ctk.CTkLabel(
            input_card,
            text="Command File",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=1, column=0, padx=(16, 8), pady=(8, 4), sticky="w")
        self.commands_file_entry = self._entry(input_card)
        self.commands_file_entry.grid(row=1, column=1, padx=(0, 10), pady=(8, 4), sticky="ew")
        self.commands_file_entry.insert(0, COMMANDS_FILE)
        self.commands_browse = ctk.CTkButton(
            input_card,
            text="Browse",
            width=100,
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(12, "bold"),
            command=self.browse_commands_file
        )
        self.commands_browse.grid(row=1, column=2, padx=(0, 16), pady=(8, 4))
        # Quick DIG
        ctk.CTkLabel(
            input_card,
            text="Quick DIG / FQDN",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="nw"
        ).grid(row=2, column=0, padx=(16, 8), pady=4, sticky="nw")
        self.quick_commands_text = ctk.CTkTextbox(
            input_card,
            height=130,
            corner_radius=10,
            fg_color=COLOR_BG,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="none"
        )
        self.quick_commands_text.grid(row=2, column=1, padx=(0, 10), pady=4, sticky="ew")
        trailing_option_frame = ctk.CTkFrame(
            input_card,
            fg_color="transparent"
        )
        trailing_option_frame.grid(
            row=2,
            column=2,
            padx=(0, 16),
            pady=4,
            sticky="nw"
        )
        ctk.CTkLabel(
            trailing_option_frame,
            text="Trailing Option",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).pack(anchor="w", pady=(0, 4))
        self.trailing_option_entry = self._entry(
            trailing_option_frame,
            width=210
        )
        self.trailing_option_entry.pack(
            fill="x",
            pady=(0, 8)
        )
        ctk.CTkLabel(
            trailing_option_frame,
            text="Empty = no trailing option",
            font=font(10),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).pack(anchor="w")
        # Output Folder
        ctk.CTkLabel(
            input_card,
            text="Output Folder",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=3, column=0, padx=(16, 8), pady=(4, 12), sticky="w")
        self.output_entry = self._entry(input_card)
        self.output_entry.grid(row=3, column=1, padx=(0, 10), pady=(4, 12), sticky="ew")
        self.output_entry.insert(0, OUTPUT_DIR)
        ctk.CTkButton(
            input_card,
            text="Browse",
            width=100,
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(12, "bold"),
            command=self.browse_output
        ).grid(row=3, column=2, padx=(0, 16), pady=(4, 12))
        # ====================================================
        # STATUS STRIP
        # ====================================================
        status_card = ModernCard(self.content, title="STATUS")
        status_card.grid(row=2, column=0, sticky="ew", padx=22, pady=10)
        status_card.grid_columnconfigure(0, weight=1)
        status_card.grid_columnconfigure(1, weight=1)
        status_card.grid_columnconfigure(2, weight=1)
        status_card.grid_columnconfigure(3, weight=1)
        self.step1_pill = StatusPill(status_card, "STEP 1", "Waiting", COLOR_TEXT_MUTED)
        self.step1_pill.grid(row=1, column=0, padx=(16, 6), pady=(4, 6), sticky="ew")
        self.step2_pill = StatusPill(status_card, "STEP 2", "Waiting", COLOR_TEXT_MUTED)
        self.step2_pill.grid(row=1, column=1, padx=6, pady=(4, 6), sticky="ew")
        self.step3_pill = StatusPill(status_card, "STEP 3", "Waiting", COLOR_TEXT_MUTED)
        self.step3_pill.grid(row=1, column=2, padx=6, pady=(4, 6), sticky="ew")
        self.excel_pill = StatusPill(status_card, "EXCEL", "Waiting", COLOR_TEXT_MUTED)
        self.excel_pill.grid(row=1, column=3, padx=(6, 16), pady=(4, 6), sticky="ew")
        # Backward-compatible references used by run_all()
        self.step1_status = self.step1_pill.text_label
        self.step2_status = self.step2_pill.text_label
        self.step3_status = self.step3_pill.text_label
        self.excel_status = self.excel_pill.text_label
        # Progress indicators
        self.step1_progress = ctk.CTkLabel(
            status_card,
            text="DIGs: 0 / 0",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.step1_progress.grid(row=2, column=0, columnspan=2, padx=(22, 6), pady=(0, 12), sticky="w")
        self.step3_progress = ctk.CTkLabel(
            status_card,
            text="DIGs: 0 / 0",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.step3_progress.grid(row=2, column=2, columnspan=2, padx=(6, 22), pady=(0, 12), sticky="w")
        # ====================================================
        # ACTION BUTTONS
        # ====================================================
        action_card = ModernCard(self.content)
        action_card.grid(row=3, column=0, sticky="ew", padx=22, pady=(10, 18))
        action_card.grid_columnconfigure(0, weight=1)
        action_card.grid_columnconfigure(1, weight=1)
        action_card.grid_columnconfigure(2, weight=1)
        self.run_button = ctk.CTkButton(
            action_card,
            text="▶   RUN ALL STEPS",
            height=48,
            corner_radius=12,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#ffffff",
            font=font(14, "bold"),
            command=self.start_run
        )
        self.run_button.grid(row=0, column=0, padx=(16, 6), pady=16, sticky="ew")
        self.send_q_button = ctk.CTkButton(
            action_card,
            text="Q  /  CONTINUE",
            height=48,
            corner_radius=12,
            fg_color=COLOR_WARNING,
            hover_color="#d97706",
            text_color="#1a1a1a",
            font=font(13, "bold"),
            command=self.send_pager_q
        )
        self.send_q_button.grid(row=0, column=1, padx=6, pady=16, sticky="ew")
        self.profiles_button = ctk.CTkButton(
            action_card,
            text="⚙   SERVER PROFILES / MULTI TEST",
            height=48,
            corner_radius=12,
            fg_color=COLOR_PURPLE,
            hover_color="#7c3aed",
            text_color="#ffffff",
            font=font(13, "bold"),
            command=self.open_server_profiles
        )
        self.profiles_button.grid(row=0, column=2, padx=(6, 16), pady=16, sticky="ew")
        # Initial authentication layout
        self.authentication_changed("Password")
        # Responsive layout
        self.root.bind("<Configure>", self._responsive_layout)
        self.root.after(150, self._responsive_layout)
    # --------------------------------------------------------
    # Entry helper
    # --------------------------------------------------------
    def _entry(self, master, width=240, show=None):
        return ctk.CTkEntry(
            master,
            height=38,
            corner_radius=10,
            fg_color=COLOR_BG,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            placeholder_text_color=COLOR_TEXT_MUTED,
            font=font(12),
            width=width,
            show=show
        )
    def _responsive_layout(self, event=None):
        """Debounced responsive layout. Applies changes only after the
        window stops moving/resizing for a short moment, and only when
        the relevant dimensions actually changed."""
        try:
            # Ignore events from child widgets — only handle the root window.
            if event is not None and event.widget is not self.root:
                return
            # Cancel any pending resize job and schedule a new one.
            if self._resize_job is not None:
                try:
                    self.root.after_cancel(self._resize_job)
                except Exception:
                    pass
            self._resize_job = self.root.after(120, self._apply_responsive_layout)
        except Exception:
            pass
    def _apply_responsive_layout(self):
        """Actually apply the responsive layout. Called after the debounce delay."""
        self._resize_job = None
        try:
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            if width < 100 or height < 100:
                return
            # Only touch widgets when the width really changed.
            if width == self._last_responsive_width:
                return
            # Quick DIG width scales with window (only reconfigure on change).
            quick_width = max(280, min(850, int(width * 0.52)))
            if quick_width != self._last_quick_width:
                try:
                    self.quick_commands_text.configure(width=quick_width)
                    self._last_quick_width = quick_width
                except Exception:
                    pass
            self._last_responsive_width = width
        except Exception:
            pass
    # ========================================================
    # FQDN SEARCH TOOL - ADDITIVE FEATURE
    # ========================================================
    def build_fqdn_tab(self):
        """Build the integrated FQDN Search Tool without changing the existing DIG flow."""
        tab = self.tabview.tab("FQDN Search Tool")
        tab.grid_rowconfigure(0, weight=0)
        tab.grid_rowconfigure(1, weight=1)
        tab.grid_columnconfigure(0, weight=1)


        self.fqdn_server_var = ctk.StringVar(value="Current Connection")
        self.fqdn_server_profiles = []
        self.fqdn_results_df = pd.DataFrame()
        self.fqdn_current_fqdn = ""
        # Display-only state. This does not change the existing search/DIG logic.
        self.fqdn_visible_columns = []
        self.fqdn_visible_rows = []
        self.fqdn_table = None


        # ----------------------------------------------------
        # Make the FQDN tab content scrollable so that the
        # header card (with CSV Folder, FQDN, DIG Server rows
        # and the Select Columns / Select Rows / Show All
        # buttons) never gets clipped on a small screen.
        # ----------------------------------------------------
        outer = ctk.CTkScrollableFrame(
            tab,
            fg_color="transparent",
            corner_radius=0
        )
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(0, weight=0)


        header = ctk.CTkFrame(outer, fg_color=COLOR_SURFACE, corner_radius=14)
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(14, 10))
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header,
            text="⌕   FQDN SEARCH + DIG",
            font=font(18, "bold"),
            text_color=COLOR_TEXT,
            anchor="w"
        ).grid(row=0, column=0, columnspan=3, padx=18, pady=(14, 2), sticky="w")
        ctk.CTkLabel(
            header,
            text="Search the existing CSV data, review forwarders, then run DIG against every IP in a selected row.",
            font=font(11),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        ).grid(row=1, column=0, columnspan=3, padx=18, pady=(0, 12), sticky="w")


        ctk.CTkLabel(
            header, text="CSV Folder", font=font(11), text_color=COLOR_TEXT_MUTED
        ).grid(row=2, column=0, padx=(18, 8), pady=(4, 12), sticky="w")
        self.fqdn_folder_entry = self._entry(header)
        self.fqdn_folder_entry.grid(row=2, column=1, padx=(0, 8), pady=(4, 12), sticky="ew")
        ctk.CTkButton(
            header, text="Browse", width=100, height=38, corner_radius=10,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT, font=font(12, "bold"),
            command=self.fqdn_select_folder
        ).grid(row=2, column=2, padx=(0, 18), pady=(4, 12))


        ctk.CTkLabel(
            header, text="FQDN", font=font(11), text_color=COLOR_TEXT_MUTED
        ).grid(row=3, column=0, padx=(18, 8), pady=(4, 12), sticky="w")
        self.fqdn_search_entry = self._entry(header)
        self.fqdn_search_entry.grid(row=3, column=1, padx=(0, 8), pady=(4, 12), sticky="ew")
        self.fqdn_search_entry.bind("<Return>", lambda event: self.fqdn_search())
        ctk.CTkButton(
            header, text="⌕  SEARCH", width=120, height=38, corner_radius=10,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
            text_color="#ffffff", font=font(12, "bold"),
            command=self.fqdn_search
        ).grid(row=3, column=2, padx=(0, 18), pady=(4, 12))


        ctk.CTkLabel(
            header, text="DIG Server", font=font(11), text_color=COLOR_TEXT_MUTED
        ).grid(row=4, column=0, padx=(18, 8), pady=(4, 14), sticky="w")
        self.fqdn_server_menu = ctk.CTkOptionMenu(
            header,
            variable=self.fqdn_server_var,
            values=["Current Connection"],
            width=280,
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            button_color=COLOR_ACCENT,
            button_hover_color=COLOR_ACCENT_HOVER,
            dropdown_fg_color=COLOR_SURFACE_ALT,
            dropdown_hover_color=COLOR_ACCENT,
            font=font(12),
            command=self.fqdn_server_changed
        )
        self.fqdn_server_menu.grid(row=4, column=1, padx=(0, 8), pady=(4, 14), sticky="w")
        ctk.CTkButton(
            header, text="↻  REFRESH SERVERS", width=160, height=38, corner_radius=10,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT, font=font(11, "bold"),
            command=self.refresh_fqdn_servers
        ).grid(row=4, column=2, padx=(0, 18), pady=(4, 14))


        results_card = ctk.CTkFrame(outer, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        results_card.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 18))
        outer.grid_rowconfigure(1, weight=1)
        results_card.grid_rowconfigure(3, weight=1)
        results_card.grid_columnconfigure(0, weight=1)


        self.fqdn_result_summary = ctk.CTkLabel(
            results_card,
            text="No search performed.",
            font=font(11, "bold"),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.fqdn_result_summary.grid(row=1, column=0, padx=16, pady=(10, 6), sticky="ew")


        # Display controls: these only affect what is shown in the table.
        # Use a CTkFrame with pack so the buttons are always visible and
        # wrap gracefully when the window is small.
        display_bar = ctk.CTkFrame(results_card, fg_color="transparent")
        display_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        display_bar.grid_columnconfigure(0, weight=1)
        display_bar.grid_columnconfigure(1, weight=0)
        display_bar.grid_columnconfigure(2, weight=0)
        display_bar.grid_columnconfigure(3, weight=0)


        self.fqdn_display_info = ctk.CTkLabel(
            display_bar,
            text="Columns: all    •    Rows: all",
            font=font(10),
            text_color=COLOR_TEXT_MUTED,
            anchor="w"
        )
        self.fqdn_display_info.grid(row=0, column=0, sticky="w", padx=(4, 8))


        self.fqdn_select_columns_button = ctk.CTkButton(
            display_bar,
            text="☷  SELECT COLUMNS",
            width=150,
            height=34,
            corner_radius=9,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(10, "bold"),
            command=self.fqdn_select_columns
        )
        self.fqdn_select_columns_button.grid(row=0, column=1, padx=4)


        self.fqdn_select_rows_button = ctk.CTkButton(
            display_bar,
            text="☑  SELECT ROWS",
            width=135,
            height=34,
            corner_radius=9,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(10, "bold"),
            command=self.fqdn_select_rows
        )
        self.fqdn_select_rows_button.grid(row=0, column=2, padx=4)


        self.fqdn_show_all_button = ctk.CTkButton(
            display_bar,
            text="SHOW ALL",
            width=95,
            height=34,
            corner_radius=9,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(10, "bold"),
            command=self.fqdn_show_all
        )
        self.fqdn_show_all_button.grid(row=0, column=3, padx=(4, 0))


        table_frame = ctk.CTkFrame(
            results_card,
            fg_color=COLOR_BG,
            corner_radius=10
        )
        table_frame.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 12))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)


        # A real Treeview table makes side-by-side row comparison much easier.
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "FQDN.Treeview",
            background=COLOR_BG,
            foreground=COLOR_TEXT,
            fieldbackground=COLOR_BG,
            rowheight=34,
            font=(FONT_FAMILY, 10)
        )
        style.configure(
            "FQDN.Treeview.Heading",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_TEXT,
            font=(FONT_FAMILY, 10, "bold"),
            relief="flat"
        )
        style.map(
            "FQDN.Treeview",
            background=[("selected", COLOR_ACCENT)],
            foreground=[("selected", "#ffffff")]
        )


        self.fqdn_table = ttk.Treeview(
            table_frame,
            show="headings",
            style="FQDN.Treeview",
            selectmode="extended"
        )
        self.fqdn_table.grid(row=0, column=0, sticky="nsew")
        self.fqdn_table.bind("<Button-1>", self.fqdn_table_click)


        y_scroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.fqdn_table.yview
        )
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.fqdn_table.xview
        )
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.fqdn_table.configure(
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set
        )


        bottom = ctk.CTkFrame(results_card, fg_color="transparent")
        bottom.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 12))
        bottom.grid_columnconfigure(0, weight=1)
        self.fqdn_export_button = ctk.CTkButton(
            bottom,
            text="⇩  EXPORT RESULTS",
            width=160,
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(11, "bold"),
            command=self.fqdn_export
        )
        self.fqdn_export_button.grid(row=0, column=1, padx=(8, 0), sticky="e")


        self.refresh_fqdn_servers()


        # Responsive reflow for the display bar buttons
        try:
            self.tabview.tab("FQDN Search Tool").bind(
                "<Configure>",
                self._fqdn_tab_configure
            )
        except Exception:
            pass


    def _fqdn_tab_configure(self, event=None):
        """Debounced handler that re-flows the FQDN display bar buttons
        so SELECT COLUMNS / SELECT ROWS / SHOW ALL stay visible when
        the FQDN tab is narrow."""
        try:
            if event is not None and event.widget is not self.tabview.tab("FQDN Search Tool"):
                return
            if self._fqdn_resize_job is not None:
                try:
                    self.root.after_cancel(self._fqdn_resize_job)
                except Exception:
                    pass
            self._fqdn_resize_job = self.root.after(120, self._apply_fqdn_reflow)
        except Exception:
            pass


    def _apply_fqdn_reflow(self):
        self._fqdn_resize_job = None
        try:
            width = self.tabview.tab("FQDN Search Tool").winfo_width()
            if width < 100:
                return
            if width == self._fqdn_last_width:
                return
            self._fqdn_last_width = width


            display_bar = self.fqdn_select_columns_button.master
            info_label = self.fqdn_display_info
            b1 = self.fqdn_select_columns_button
            b2 = self.fqdn_select_rows_button
            b3 = self.fqdn_show_all_button


            # Reset everything to a single-row layout first.
            for widget in (info_label, b1, b2, b3):
                try:
                    widget.grid_forget()
                except Exception:
                    pass


            # Wide layout: info + three buttons in a single row.
            if width >= 720:
                display_bar.grid_columnconfigure(0, weight=1)
                display_bar.grid_columnconfigure(1, weight=0)
                display_bar.grid_columnconfigure(2, weight=0)
                display_bar.grid_columnconfigure(3, weight=0)
                info_label.grid(row=0, column=0, sticky="w", padx=(4, 8))
                b1.grid(row=0, column=1, padx=4)
                b2.grid(row=0, column=2, padx=4)
                b3.grid(row=0, column=3, padx=(4, 0))
                return


            # Narrow layout: info on its own row, then the three buttons
            # on the next row so nothing gets clipped.
            display_bar.grid_columnconfigure(0, weight=1)
            display_bar.grid_columnconfigure(1, weight=0)
            display_bar.grid_columnconfigure(2, weight=0)
            display_bar.grid_columnconfigure(3, weight=0)
            info_label.grid(row=0, column=0, columnspan=4, sticky="w", padx=(4, 8))
            b1.grid(row=1, column=0, padx=4, pady=(6, 0), sticky="w")
            b2.grid(row=1, column=1, padx=4, pady=(6, 0), sticky="w")
            b3.grid(row=1, column=2, padx=(4, 0), pady=(6, 0), sticky="w")
        except Exception:
            pass


    def fqdn_update_display_info(self):
        total_cols = len(self.fqdn_results_df.columns) if not self.fqdn_results_df.empty else 0
        visible_cols = len(self.fqdn_visible_columns)
        total_rows = len(self.fqdn_results_df.index) if not self.fqdn_results_df.empty else 0
        visible_rows = len(self.fqdn_visible_rows)
        if total_cols:
            col_text = f"{visible_cols}/{total_cols}"
        else:
            col_text = "0"
        if total_rows:
            row_text = f"{visible_rows}/{total_rows}"
        else:
            row_text = "0"
        try:
            self.fqdn_display_info.configure(
                text=f"Columns: {col_text}    •    Rows: {row_text}    •    DIG is available in the last column"
            )
        except Exception:
            pass


    def fqdn_show_all(self):
        if self.fqdn_results_df.empty:
            return
        self.fqdn_visible_columns = list(self.fqdn_results_df.columns)
        self.fqdn_visible_rows = list(range(len(self.fqdn_results_df.index)))
        self.fqdn_render_table()
        self.fqdn_update_display_info()


    def fqdn_select_columns(self):
        if self.fqdn_results_df.empty:
            messagebox.showinfo("FQDN Display", "Search for an FQDN first.")
            return


        window = ctk.CTkToplevel(self.root)
        window.title("Select Columns to Display")
        window.geometry("560x620")
        window.minsize(480, 480)
        window.transient(self.root)
        window.configure(fg_color=COLOR_BG)


        ctk.CTkLabel(
            window,
            text="SELECT COLUMNS",
            font=font(16, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=18, pady=(16, 2))
        ctk.CTkLabel(
            window,
            text="Choose which FQDN result columns should be visible in the comparison table.",
            font=font(10),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=18, pady=(0, 10))


        selected = {
            column: ctk.BooleanVar(value=column in self.fqdn_visible_columns)
            for column in self.fqdn_results_df.columns
        }


        scroll = ctk.CTkScrollableFrame(
            window, fg_color=COLOR_SURFACE, corner_radius=12
        )
        scroll.pack(fill="both", expand=True, padx=16, pady=8)


        for index, column in enumerate(self.fqdn_results_df.columns):
            ctk.CTkCheckBox(
                scroll,
                text=str(column),
                variable=selected[column],
                font=font(11),
                text_color=COLOR_TEXT
            ).grid(row=index, column=0, padx=14, pady=7, sticky="w")


        buttons = ctk.CTkFrame(window, fg_color="transparent")
        buttons.pack(fill="x", padx=16, pady=(6, 16))


        def select_all():
            for var in selected.values():
                var.set(True)


        def clear_all():
            for var in selected.values():
                var.set(False)


        def apply():
            chosen = [
                column for column in self.fqdn_results_df.columns
                if selected[column].get()
            ]
            if not chosen:
                messagebox.showwarning(
                    "FQDN Display",
                    "Select at least one column."
                )
                return
            self.fqdn_visible_columns = chosen
            self.fqdn_render_table()
            self.fqdn_update_display_info()
            window.destroy()


        ctk.CTkButton(
            buttons, text="ALL", width=80, height=36,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            command=select_all
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="CLEAR", width=80, height=36,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            command=clear_all
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="APPLY", width=100, height=36,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
            text_color="#ffffff", font=font(11, "bold"),
            command=apply
        ).pack(side="right", padx=4)


    def fqdn_select_rows(self):
        if self.fqdn_results_df.empty:
            messagebox.showinfo("FQDN Display", "Search for an FQDN first.")
            return


        window = ctk.CTkToplevel(self.root)
        window.title("Select Rows to Display")
        window.geometry("700x620")
        window.minsize(560, 480)
        window.transient(self.root)
        window.configure(fg_color=COLOR_BG)


        ctk.CTkLabel(
            window,
            text="SELECT ROWS",
            font=font(16, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=18, pady=(16, 2))
        ctk.CTkLabel(
            window,
            text="Select the result rows you want to keep visible in the comparison table.",
            font=font(10),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=18, pady=(0, 10))


        selected = {
            index: ctk.BooleanVar(value=index in self.fqdn_visible_rows)
            for index in range(len(self.fqdn_results_df.index))
        }


        scroll = ctk.CTkScrollableFrame(
            window, fg_color=COLOR_SURFACE, corner_radius=12
        )
        scroll.pack(fill="both", expand=True, padx=16, pady=8)


        columns = list(self.fqdn_results_df.columns)
        preview_columns = columns[:4]


        for index, row in self.fqdn_results_df.iterrows():
            preview_parts = []
            for column in preview_columns:
                value = "" if pd.isna(row[column]) else str(row[column])
                if len(value) > 80:
                    value = value[:77] + "..."
                preview_parts.append(f"{column}: {value}")
            preview = " | ".join(preview_parts)


            ctk.CTkCheckBox(
                scroll,
                text=f"ROW {index + 1}   {preview}",
                variable=selected[index],
                font=font(10),
                text_color=COLOR_TEXT
            ).grid(row=index, column=0, padx=14, pady=7, sticky="w")


        buttons = ctk.CTkFrame(window, fg_color="transparent")
        buttons.pack(fill="x", padx=16, pady=(6, 16))


        def select_all():
            for var in selected.values():
                var.set(True)


        def clear_all():
            for var in selected.values():
                var.set(False)


        def apply():
            chosen = [
                index for index in range(len(self.fqdn_results_df.index))
                if selected[index].get()
            ]
            if not chosen:
                messagebox.showwarning(
                    "FQDN Display",
                    "Select at least one row."
                )
                return
            self.fqdn_visible_rows = chosen
            self.fqdn_render_table()
            self.fqdn_update_display_info()
            window.destroy()


        ctk.CTkButton(
            buttons, text="ALL", width=80, height=36,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            command=select_all
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="CLEAR", width=80, height=36,
            fg_color=COLOR_SURFACE_ALT, hover_color=COLOR_BORDER,
            command=clear_all
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            buttons, text="APPLY", width=100, height=36,
            fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER,
            text_color="#ffffff", font=font(11, "bold"),
            command=apply
        ).pack(side="right", padx=4)


    def fqdn_render_table(self):
        if self.fqdn_table is None:
            return


        for item in self.fqdn_table.get_children():
            self.fqdn_table.delete(item)


        if self.fqdn_results_df.empty:
            self.fqdn_table["columns"] = ()
            return


        columns = list(self.fqdn_visible_columns or self.fqdn_results_df.columns)
        table_columns = [f"c{index}" for index in range(len(columns))] + ["__dig__"]
        self.fqdn_table["columns"] = table_columns


        for index, column in enumerate(columns):
            table_column = f"c{index}"
            self.fqdn_table.heading(table_column, text=str(column))
            # Wider columns for forwarder/zone-like data, compact for others.
            sample_values = [
                "" if pd.isna(v) else str(v)
                for v in self.fqdn_results_df[column].tolist()
            ]
            max_len = max([len(str(column))] + [len(v) for v in sample_values[:100]])
            width = min(600, max(110, max_len * 7 + 24))
            if "allow" in str(column).lower():
                width = max(110, min(width, 150))
            self.fqdn_table.column(
                table_column,
                width=width,
                minwidth=90,
                anchor="w",
                stretch=False
            )


        self.fqdn_table.heading("__dig__", text="ACTION")
        self.fqdn_table.column(
            "__dig__", width=110, minwidth=110, anchor="center", stretch=False
        )


        for dataframe_index in self.fqdn_visible_rows:
            if dataframe_index >= len(self.fqdn_results_df.index):
                continue
            row = self.fqdn_results_df.iloc[dataframe_index]
            values = [
                "" if pd.isna(row[column]) else str(row[column])
                for column in columns
            ]
            # DIG still uses the complete original row, including hidden columns.
            all_row_values = [
                "" if pd.isna(value) else str(value)
                for value in row.tolist()
            ]
            ips = self.fqdn_extract_ips(all_row_values)
            action_text = f"▶ DIG ({len(ips)})" if ips else "NO IP"
            item_id = self.fqdn_table.insert(
                "",
                "end",
                iid=f"row_{dataframe_index}",
                values=values + [action_text]
            )
    
    def fqdn_table_click(self, event):
        """Run the existing row-DIG logic when the ACTION/DIG cell is clicked."""
        if self.fqdn_table is None:
            return


        region = self.fqdn_table.identify("region", event.x, event.y)
        if region != "cell":
            return


        column_id = self.fqdn_table.identify_column(event.x)
        row_id = self.fqdn_table.identify_row(event.y)
        if not row_id or column_id != f"#{len(self.fqdn_visible_columns) + 1}":
            return


        try:
            dataframe_index = int(row_id.replace("row_", ""))
            row = self.fqdn_results_df.iloc[dataframe_index]
            row_values = [
                "" if pd.isna(value) else str(value)
                for value in row.tolist()
            ]
            ips = self.fqdn_extract_ips(row_values)
            if ips:
                self.fqdn_run_row_dig(row_values, ips)
        except Exception as exc:
            messagebox.showerror("FQDN DIG", f"Unable to run row DIG:\n{exc}")


    def fqdn_select_folder(self):
        folder = filedialog.askdirectory(title="Select CSV Folder")
        if folder:
            self.fqdn_folder_entry.delete(0, "end")
            self.fqdn_folder_entry.insert(0, folder)


    def refresh_fqdn_servers(self):
        profiles = self.load_server_profiles()
        self.fqdn_server_profiles = profiles
        values = ["Current Connection"]
        for profile in profiles:
            name = profile.get("name", "").strip()
            server = profile.get("server", "").strip()
            label = f"{name}  ({server})" if name and server else (name or server)
            if label and label not in values:
                values.append(label)
        self.fqdn_server_menu.configure(values=values)
        if self.fqdn_server_var.get() not in values:
            self.fqdn_server_var.set("Current Connection")


    def fqdn_server_changed(self, value):
        if value == "Current Connection":
            return
        profile = self.fqdn_profile_from_selection(value)
        if not profile:
            return
        # Apply the selected profile to the existing connection fields.
        # Password is intentionally not stored by the existing profile logic,
        # so preserve the currently entered password.
        current_password = self.password_entry.get()
        self.server_entry.delete(0, "end")
        self.server_entry.insert(0, profile.get("server", ""))
        self.username_entry.delete(0, "end")
        self.username_entry.insert(0, profile.get("username", ""))
        self.port_entry.delete(0, "end")
        self.port_entry.insert(0, profile.get("port", "22"))
        self.auth_var.set(profile.get("authentication", "Password"))
        self.authentication_changed(self.auth_var.get())
        self.password_entry.delete(0, "end")
        self.password_entry.insert(0, current_password)
        self.key_entry.delete(0, "end")
        self.key_entry.insert(0, profile.get("key_file", ""))
        self.key_passphrase_entry.delete(0, "end")
        self.key_passphrase_entry.insert(0, profile.get("key_passphrase", ""))
        self.log(f"[FQDN] Selected DIG server profile: {profile.get('name', value)}")


    def fqdn_profile_from_selection(self, value):
        for profile in self.fqdn_server_profiles:
            name = profile.get("name", "").strip()
            server = profile.get("server", "").strip()
            label = f"{name}  ({server})" if name and server else (name or server)
            if label == value:
                return profile
        return None


    def fqdn_search(self):
        fqdn = self.fqdn_search_entry.get().strip()
        folder_path = self.fqdn_folder_entry.get().strip()


        if not folder_path:
            messagebox.showerror("FQDN Search", "Please select folder first")
            return
        if not fqdn:
            messagebox.showerror("FQDN Search", "Enter FQDN")
            return
        if not os.path.isdir(folder_path):
            messagebox.showerror("FQDN Search", f"Folder not found:\n{folder_path}")
            return


        # Existing Search FQDN.py search logic is preserved here.
        required_columns = [0, 2, 3, 4, 6, 8, 9, 10, 27, 30, 37, 39, 40]
        all_results = []
        for file in os.listdir(folder_path):
            if file.endswith(".csv"):
                file_path = os.path.join(folder_path, file)
                try:
                    df = pd.read_csv(file_path, dtype=str, low_memory=False)
                    if len(df.columns) < 2:
                        continue
                    filtered = df[df.iloc[:, 1] == fqdn]
                    if not filtered.empty:
                        selected = filtered.iloc[:, required_columns]
                        selected["Source_File"] = file
                        all_results.append(selected)
                except Exception as e:
                    print(f"Error reading {file}: {e}")


        if all_results:
            self.fqdn_results_df = pd.concat(all_results, ignore_index=True)
            self.fqdn_current_fqdn = fqdn
            self.fqdn_show_results()
        else:
            self.fqdn_results_df = pd.DataFrame()
            self.fqdn_current_fqdn = fqdn
            self.fqdn_clear_results()
            messagebox.showinfo("FQDN Search", "No match found")


    def fqdn_clear_results(self):
        if self.fqdn_table is not None:
            for item in self.fqdn_table.get_children():
                self.fqdn_table.delete(item)
            self.fqdn_table["columns"] = ()
        self.fqdn_visible_columns = []
        self.fqdn_visible_rows = []
        self.fqdn_result_summary.configure(text="No matching records found.")
        self.fqdn_update_display_info()


    def fqdn_extract_ips(self, row_values):
        """Extract all IPv4 addresses present in a search-result row."""
        ipv4_pattern = r"\b(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}\b"
        ips = []
        for value in row_values:
            if value is None:
                continue
            text_value = str(value)
            for ip in re.findall(ipv4_pattern, text_value):
                if ip not in ips:
                    ips.append(ip)
        return ips


    def fqdn_show_results(self):
        # Display-only changes: keep all columns/rows selected after each search.
        self.fqdn_visible_columns = list(self.fqdn_results_df.columns)
        self.fqdn_visible_rows = list(range(len(self.fqdn_results_df.index)))
        self.fqdn_render_table()


        count = len(self.fqdn_results_df.index)
        self.fqdn_result_summary.configure(
            text=(
                f"FQDN: {self.fqdn_current_fqdn}    •    Matching rows: {count}    "
                f"•    Select columns/rows for comparison or click ▶ DIG in the ACTION column."
            )
        )
        self.fqdn_update_display_info()


    def fqdn_run_row_dig(self, row_values, ips):
        if not ips:
            return
        fqdn = self.fqdn_current_fqdn.strip()
        if not fqdn:
            messagebox.showerror("FQDN DIG", "No FQDN selected.")
            return


        try:
            server = self.server_entry.get().strip()
            username = self.username_entry.get().strip()
            port = self.port_entry.get().strip()
            authentication = self.auth_var.get()
            password = self.password_entry.get()
            key_file = self.key_entry.get().strip()
            key_passphrase = self.key_passphrase_entry.get()


            if not server or not username:
                raise Exception("Select a DIG server first.")
            if authentication == "Password" and not password:
                raise Exception("SSH password is required for the selected DIG server.")
            if authentication == "SSH Key" and not key_file:
                raise Exception("Please select a private key for the selected DIG server.")
            if authentication == "SSH Key" and not os.path.exists(key_file):
                raise Exception("Selected key does not exist.")
        except Exception as exc:
            messagebox.showerror("FQDN DIG", str(exc))
            return


        commands = [f"dig @{ip} {fqdn}" for ip in ips]
        self.open_live_output()
        self.log("\n" + "=" * 70)
        self.log("FQDN SEARCH TOOL - ROW DIG")
        self.log("=" * 70)
        self.log(f"FQDN    : {fqdn}")
        self.log(f"Server  : {server}")
        self.log(f"Commands: {len(commands)}")
        for command in commands:
            self.log(f"[FQDN DIG] {command}")


        def worker():
            try:
                output = execute_ssh(
                    server=server,
                    username=username,
                    port=port,
                    authentication=authentication,
                    password=password,
                    key_file=key_file,
                    key_passphrase=key_passphrase,
                    commands=commands,
                    log_callback=self.log,
                    keep_connection=False
                )
                if output:
                    self.log("[FQDN DIG] Row DIG completed.")
                else:
                    self.log("[FQDN DIG] No output returned.")
            except Exception as exc:
                self.log(f"[FQDN DIG] ERROR: {exc}")


        threading.Thread(target=worker, daemon=True).start()


    def fqdn_export(self):
        if self.fqdn_results_df.empty:
            messagebox.showerror("FQDN Search", "No data to export")
            return
        file_path = filedialog.asksaveasfilename(
            title="Export FQDN Search Results",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.fqdn_results_df.to_csv(file_path, index=False)
            messagebox.showinfo("FQDN Search", f"File saved: {file_path}")


    def authentication_changed(self, value):
        if value == "Password":
            self.key_label.grid_remove()
            self.key_entry.grid_remove()
            self.key_browse.grid_remove()
            self.key_passphrase_label.grid_remove()
            self.key_passphrase_entry.grid_remove()
            self.password_label.grid()
            self.password_entry.grid()
        else:
            self.password_label.grid_remove()
            self.password_entry.grid_remove()
            self.key_label.grid(
                row=3,
                column=0,
                padx=(16, 8),
                pady=4,
                sticky="w"
            )
            self.key_entry.grid(
                row=3,
                column=1,
                padx=(0, 10),
                pady=4,
                sticky="ew"
            )
            self.key_browse.grid(
                row=3,
                column=2,
                padx=(0, 16),
                pady=4,
                sticky="w"
            )
            self.key_passphrase_label.grid(
                row=4,
                column=0,
                padx=(16, 8),
                pady=4,
                sticky="w"
            )
            self.key_passphrase_entry.grid(
                row=4,
                column=1,
                columnspan=3,
                padx=(0, 16),
                pady=4,
                sticky="ew"
            )
    # ========================================================
    # BROWSE COMMAND FILE
    # ========================================================
    def browse_commands_file(self):
        file_path = filedialog.askopenfilename(
            title="Select DIG Command File",
            filetypes=[
                ("All Files", "*.*"),
                ("Text Files", "*.txt"),
                ("Excel Files", "*.xlsx *.xlsm"),
                ("CSV Files", "*.csv")
            ]
        )
        if file_path:
            self.commands_file_entry.delete(0, "end")
            self.commands_file_entry.insert(0, file_path)
    # ========================================================
    # BROWSE KEY
    # ========================================================
    def browse_key(self):
        file_path = filedialog.askopenfilename(
            title="Select SSH Private Key",
            filetypes=[
                ("All Files", "*.*"),
                ("SSH Private Keys", "*.pem *.ppk"),
                ("PEM Files", "*.pem"),
                ("PuTTY Private Keys", "*.ppk")
            ]
        )
        if file_path:
            self.key_entry.delete(0, "end")
            self.key_entry.insert(0, file_path)
    # ========================================================
    # EFFECTIVE COMMAND INPUT
    # ========================================================
    def get_effective_commands(self):
        """Use pasted DIG/FQDN input when present; otherwise use the existing command file."""
        quick_input = self.quick_commands_text.get(
            "1.0",
            "end"
        ).strip()
        if not quick_input:
            commands_file = (
                self.commands_file_entry.get()
                .strip()
            )
            return read_commands(commands_file, self.trailing_option_entry.get().strip())
        # Create a temporary TXT command file so the existing
        # read_commands() logic remains completely unchanged.
        #
        # Quick input supports:
        #
        # 1. Normal DIG command:
        #       dig @10.10.10.10 example.com
        #
        # 2. FQDN only:
        #       example.com
        #
        # 3. FQDN + forward_to format (TAB separated):
        #       example.com    ns1.example.com/10.10.10.10,ns2.example.com/10.10.10.20
        #
        # The third format is expanded to individual DIG commands.
        commands_for_temp_file = []
        for raw_line in quick_input.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            # ------------------------------------------------
            # FQDN + forward_to input
            # ------------------------------------------------
            # The source format uses TAB between the FQDN and
            # the forward_to column.
            # ------------------------------------------------
            if "\t" in raw_line:
                parts = raw_line.split("\t", 1)
                fqdn = parts[0].strip()
                forward_to = parts[1].strip()
                if fqdn and forward_to:
                    for target in forward_to.split(","):
                        target = target.strip()
                        if "/" not in target:
                            continue
                        ip = target.rsplit(
                            "/",
                            1
                        )[1].strip()
                        if not ip:
                            continue
                        trailing_option = self.trailing_option_entry.get().strip()
                        command = f"dig @{ip} {fqdn}"
                        if trailing_option:
                            command += f" {trailing_option}"
                        commands_for_temp_file.append(command)
                    continue
            # ------------------------------------------------
            # Existing Quick DIG command / FQDN behavior
            # ------------------------------------------------
            commands_for_temp_file.append(line)
        if not commands_for_temp_file:
            return []
        temp_path = None
        try:
            fd, temp_path = tempfile.mkstemp(
                prefix="infoblox_quick_commands_",
                suffix=".txt"
            )
            with os.fdopen(
                fd,
                "w",
                encoding="utf-8"
            ) as temp_file:
                for command in commands_for_temp_file:
                    temp_file.write(command)
                    temp_file.write("\n")
            self.log(
                "[COMMAND INPUT] Using commands/FQDN from Quick DIG input box."
            )
            return read_commands(temp_path, self.trailing_option_entry.get().strip())
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
    # ========================================================
    # OUTPUT FOLDER
    # ========================================================
    def browse_output(self):
        folder = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if folder:
            self.output_entry.delete(
                0,
                "end"
            )
            self.output_entry.insert(
                0,
                folder
            )
    # ========================================================
    # LIVE OUTPUT WINDOW
    # ========================================================
    # ========================================================
    # INTERACTIVE SSH TERMINAL
    # ========================================================
    def open_ssh_terminal(self):
        """Open the interactive SSH terminal window."""
        try:
            if self.terminal_window is not None and self.terminal_window.winfo_exists():
                self.terminal_window.deiconify()
                self.terminal_window.lift()
                self.terminal_window.focus_force()
                return
            with self.terminal_output_lock:
                self.terminal_output_buffer = ""
            self.terminal_window = ctk.CTkToplevel(self.root)
            self.terminal_window.title("INFOBLOX — Interactive SSH Terminal")
            self.terminal_window.geometry("1100x700")
            self.terminal_window.minsize(750, 450)
            self.terminal_window.configure(fg_color=COLOR_BG)
            self.terminal_window.protocol("WM_DELETE_WINDOW", self.close_ssh_terminal)
            self.terminal_window.grid_rowconfigure(1, weight=1)
            self.terminal_window.grid_columnconfigure(0, weight=1)
            header = ctk.CTkFrame(
                self.terminal_window,
                fg_color=COLOR_SURFACE,
                corner_radius=12
            )
            header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 7))
            header.grid_columnconfigure(0, weight=1)
            self.terminal_status_label = ctk.CTkLabel(
                header,
                text="●  Disconnected",
                font=font(13, "bold"),
                text_color=COLOR_TEXT_MUTED,
                anchor="w"
            )
            self.terminal_status_label.grid(
                row=0, column=0, padx=14, pady=12, sticky="w"
            )
            self.terminal_connect_button = ctk.CTkButton(
                header,
                text="CONNECT",
                width=110,
                height=34,
                corner_radius=9,
                fg_color=COLOR_SUCCESS,
                hover_color="#059669",
                text_color="#ffffff",
                font=font(12, "bold"),
                command=self.connect_ssh_terminal
            )
            self.terminal_connect_button.grid(
                row=0, column=1, padx=(6, 8), pady=10
            )
            self.terminal_disconnect_button = ctk.CTkButton(
                header,
                text="DISCONNECT",
                width=120,
                height=34,
                corner_radius=9,
                fg_color=COLOR_DANGER,
                hover_color="#dc2626",
                text_color="#ffffff",
                font=font(12, "bold"),
                command=self.disconnect_ssh_terminal,
                state="disabled"
            )
            self.terminal_disconnect_button.grid(
                row=0, column=2, padx=(0, 8), pady=10
            )
            self.terminal_save_button = ctk.CTkButton(
                header,
                text="SAVE TXT",
                width=100,
                height=34,
                corner_radius=9,
                fg_color=COLOR_SURFACE_ALT,
                hover_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=font(12, "bold"),
                command=self.save_ssh_terminal_output
            )
            self.terminal_save_button.grid(
                row=0, column=3, padx=(0, 14), pady=10
            )
            self.terminal_text = ctk.CTkTextbox(
                self.terminal_window,
                corner_radius=10,
                fg_color="#080b0f",
                border_width=1,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=ctk.CTkFont(family="Consolas", size=12),
                wrap="none"
            )
            self.terminal_text.grid(
                row=1, column=0, sticky="nsew", padx=14, pady=7
            )
            command_frame = ctk.CTkFrame(
                self.terminal_window,
                fg_color=COLOR_SURFACE,
                corner_radius=12
            )
            command_frame.grid(
                row=2, column=0, sticky="ew", padx=14, pady=(7, 14)
            )
            command_frame.grid_columnconfigure(0, weight=1)
            self.terminal_command_entry = ctk.CTkEntry(
                command_frame,
                height=40,
                corner_radius=10,
                fg_color=COLOR_BG,
                border_width=1,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=ctk.CTkFont(family="Consolas", size=12),
                placeholder_text="Enter Infoblox command and press Enter..."
            )
            self.terminal_command_entry.grid(
                row=0, column=0, padx=(12, 8), pady=12, sticky="ew"
            )
            # ----------------------------------------------------
            # ENTER KEY HANDLING — command entry only
            # ----------------------------------------------------
            # The diagnostics showed that the output Text widget was
            # receiving keyboard focus. Do not bind Enter to that widget.
            # Keep Enter exclusively on the real Tk Entry and prevent the
            # read-only output Text widget from taking focus on mouse click.
            # ----------------------------------------------------
            try:
                entry_widget = self.terminal_command_entry._entry


                # Find the real Tk Text widget inside CTkTextbox.
                terminal_text_widget = None
                for child in self.terminal_text.winfo_children():
                    try:
                        if child.winfo_class() == "Text":
                            terminal_text_widget = child
                            break
                    except Exception:
                        pass


                def terminal_enter(event=None):
                    print(
                        "[SSH TERMINAL] ENTER FIRED:",
                        repr(event.widget) if event is not None else None
                    )
                    return self.send_ssh_terminal_command(event)


                # Keep the command Entry completely normal so typing is not
                # affected. Handle Enter directly on the real Tk Entry.
                entry_widget.bind("<Return>", terminal_enter, add="+")
                entry_widget.bind("<KP_Enter>", terminal_enter, add="+")


                # Also handle Enter when the terminal output Text widget has
                # focus. Do not redirect focus or intercept mouse clicks; this
                # keeps normal typing/focus behavior intact.
                if terminal_text_widget is not None:
                    terminal_text_widget.bind(
                        "<Return>",
                        terminal_enter,
                        add="+"
                    )
                    terminal_text_widget.bind(
                        "<KP_Enter>",
                        terminal_enter,
                        add="+"
                    )


                # Give the command Entry the initial keyboard focus once.
                self.terminal_window.after(100, entry_widget.focus_set)


            except Exception as exc:
                print("Enter binding failed:", exc)


            ctk.CTkButton(
                command_frame,
                text="SEND",
                width=90,
                height=40,
                corner_radius=10,
                fg_color=COLOR_ACCENT,
                hover_color=COLOR_ACCENT_HOVER,
                font=font(12, "bold"),
                command=self.send_ssh_terminal_command
            ).grid(
                row=0, column=1, padx=(0, 12), pady=12
            )
            self.terminal_write(
                "Interactive SSH Terminal\n"
                "Uses the Server / Username / Port / Authentication "
                "fields from the main window.\n\n"
            )
        except Exception as exc:
            self.log(f"[SSH TERMINAL] Unable to open terminal: {exc}")
    def terminal_write(self, text):
        """Append terminal output and keep a complete TXT-save buffer."""
        if not text:
            return
        cleaned = clean_terminal_text(text)
        with self.terminal_output_lock:
            self.terminal_output_buffer += cleaned
        def update():
            try:
                if self.terminal_text is not None and self.terminal_text.winfo_exists():
                    self.terminal_text.insert("end", cleaned)
                    self.terminal_text.see("end")
            except Exception:
                pass
        try:
            self.root.after(0, update)
        except Exception:
            pass
    def connect_ssh_terminal(self):
        """Connect an interactive SSH shell using the main connection fields."""
        if self.terminal_connected:
            self.terminal_write("\n[SSH] Already connected.\n")
            return
        server = self.server_entry.get().strip()
        username = self.username_entry.get().strip()
        port = self.port_entry.get().strip() or "22"
        authentication = self.auth_var.get()
        password = self.password_entry.get()
        key_file = self.key_entry.get().strip()
        key_passphrase = self.key_passphrase_entry.get()
        if not server or not username:
            messagebox.showwarning(
                "SSH Terminal",
                "Please enter Server / IP and Username in the main Connection section."
            )
            return
        try:
            port_value = int(port)
        except ValueError:
            messagebox.showwarning(
                "SSH Terminal",
                "SSH Port must be a valid number."
            )
            return
        self.terminal_connect_button.configure(state="disabled")
        self.terminal_status_label.configure(
            text="●  Connecting...",
            text_color=COLOR_WARNING
        )
        self.terminal_write(
            f"\n[SSH] Connecting to {server}:{port_value}...\n"
        )
        def worker():
            client = None
            temp_key_dir = None
            try:
                client = paramiko.SSHClient()
                client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                if authentication == "Password":
                    client.connect(
                        hostname=server,
                        port=port_value,
                        username=username,
                        password=password,
                        timeout=20,
                        banner_timeout=20,
                        auth_timeout=20,
                        look_for_keys=False,
                        allow_agent=False
                    )
                else:
                    private_key, temp_key_dir = load_private_key(
                        key_file,
                        key_passphrase
                    )
                    client.connect(
                        hostname=server,
                        port=port_value,
                        username=username,
                        pkey=private_key,
                        timeout=20,
                        banner_timeout=20,
                        auth_timeout=20,
                        look_for_keys=False,
                        allow_agent=False
                    )
                channel = client.invoke_shell(
                    term="xterm",
                    width=160,
                    height=200
                )
                channel.settimeout(1.0)
                self.terminal_client = client
                self.terminal_channel = channel
                self.terminal_temp_key_dir = temp_key_dir
                self.terminal_connected = True
                self.terminal_write(f"[SSH] Connected to {server}\n")
                def connected_ui():
                    try:
                        self.terminal_status_label.configure(
                            text=f"●  Connected: {server}",
                            text_color=COLOR_SUCCESS
                        )
                        self.terminal_connect_button.configure(
                            state="disabled"
                        )
                        self.terminal_disconnect_button.configure(
                            state="normal"
                        )
                        self.terminal_window.focus_force()
                        command_entry_widget = self.terminal_command_entry._entry
                        command_entry_widget.focus_force()
                        # Re-assert focus after the toplevel finishes its
                        # activation/focus processing.
                        self.terminal_window.after(100, command_entry_widget.focus_force)
                        self.terminal_window.after(300, command_entry_widget.focus_force)
                    except Exception:
                        pass
                self.root.after(0, connected_ui)
                def reader():
                    try:
                        while self.terminal_connected:
                            try:
                                if channel.recv_ready():
                                    data = channel.recv(65535)
                                    if not data:
                                        break
                                    self.terminal_write(
                                        data.decode("utf-8", errors="replace")
                                    )
                                else:
                                    time.sleep(0.05)
                            except Exception:
                                if not self.terminal_connected:
                                    break
                                time.sleep(0.1)
                    finally:
                        if self.terminal_connected:
                            self.root.after(0, self._terminal_connection_lost)
                self.terminal_reader_thread = threading.Thread(
                    target=reader,
                    daemon=True
                )
                self.terminal_reader_thread.start()
            except Exception as exc:
                if temp_key_dir:
                    shutil.rmtree(temp_key_dir, ignore_errors=True)
                try:
                    if client:
                        client.close()
                except Exception:
                    pass
                self.terminal_client = None
                self.terminal_channel = None
                self.terminal_connected = False
                self.terminal_write(f"[SSH ERROR] {exc}\n")
                def failed_ui():
                    try:
                        self.terminal_status_label.configure(
                            text="●  Connection failed",
                            text_color=COLOR_DANGER
                        )
                        self.terminal_connect_button.configure(state="normal")
                    except Exception:
                        pass
                self.root.after(0, failed_ui)
        threading.Thread(target=worker, daemon=True).start()
    def send_ssh_terminal_command(self, event=None):
        """Send the manually entered command exactly as entered."""
        if not self.terminal_connected or self.terminal_channel is None:
            self.terminal_write(
                "\n[SSH] Not connected. Click CONNECT first.\n"
            )
            return "break"
        try:
            command = self.terminal_command_entry.get()
            if not command:
                return "break"
            self.terminal_command_entry.delete(0, "end")
            self.terminal_channel.send(command + "\n")
        except Exception as exc:
            self.terminal_write(
                f"\n[SSH ERROR] Failed to send command: {exc}\n"
            )
        return "break"
    def disconnect_ssh_terminal(self):
        """Disconnect the interactive SSH session."""
        self.terminal_connected = False
        channel = self.terminal_channel
        client = self.terminal_client
        temp_key_dir = self.terminal_temp_key_dir
        self.terminal_channel = None
        self.terminal_client = None
        self.terminal_temp_key_dir = None
        try:
            if channel:
                try:
                    channel.send("exit\n")
                except Exception:
                    pass
                try:
                    channel.close()
                except Exception:
                    pass
        finally:
            try:
                if client:
                    client.close()
            except Exception:
                pass
            if temp_key_dir:
                shutil.rmtree(
                    temp_key_dir,
                    ignore_errors=True
                )
        self.terminal_write("\n[SSH] Disconnected.\n")
        try:
            self.terminal_status_label.configure(
                text="●  Disconnected",
                text_color=COLOR_TEXT_MUTED
            )
            self.terminal_connect_button.configure(state="normal")
            self.terminal_disconnect_button.configure(state="disabled")
        except Exception:
            pass
    def _terminal_connection_lost(self):
        if not self.terminal_connected:
            return
        self.terminal_connected = False
        self.terminal_channel = None
        self.terminal_client = None
        try:
            self.terminal_status_label.configure(
                text="●  Connection lost",
                text_color=COLOR_DANGER
            )
            self.terminal_connect_button.configure(state="normal")
            self.terminal_disconnect_button.configure(state="disabled")
        except Exception:
            pass
        self.terminal_write(
            "\n[SSH] Connection closed by remote server.\n"
        )
    def save_ssh_terminal_output(self):
        """Save the interactive terminal output as TXT."""
        with self.terminal_output_lock:
            output = self.terminal_output_buffer
        if not output.strip():
            messagebox.showinfo(
                "Save Terminal Output",
                "There is no terminal output to save."
            )
            return
        default_dir = self.output_entry.get().strip() or OUTPUT_DIR
        try:
            os.makedirs(default_dir, exist_ok=True)
        except Exception:
            default_dir = OUTPUT_DIR
        filename = (
            "infoblox_ssh_terminal_"
            + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".txt"
        )
        file_path = filedialog.asksaveasfilename(
            title="Save SSH Terminal Output",
            initialdir=default_dir,
            initialfile=filename,
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(output)
            self.log(f"[SSH TERMINAL] Output saved: {file_path}")
            messagebox.showinfo(
                "Save Terminal Output",
                f"Output saved successfully:\n\n{file_path}"
            )
        except Exception as exc:
            messagebox.showerror(
                "Save Terminal Output",
                f"Unable to save terminal output:\n\n{exc}"
            )
    def close_ssh_terminal(self):
        """Close the terminal window and its SSH session."""
        try:
            bindtag = getattr(self, "_terminal_enter_bindtag", None)
            if bindtag and self.terminal_window is not None:
                self.terminal_window.unbind_class(bindtag, "<Return>")
                self.terminal_window.unbind_class(bindtag, "<KP_Enter>")
            self._terminal_enter_bindtag = None
        except Exception:
            pass
        self.disconnect_ssh_terminal()
        try:
            if self.terminal_window is not None and self.terminal_window.winfo_exists():
                self.terminal_window.destroy()
        except Exception:
            pass
        self.terminal_window = None
        self.terminal_text = None
        self.terminal_command_entry = None
        self.terminal_status_label = None
        self.terminal_connect_button = None
        self.terminal_disconnect_button = None
    def open_live_output(self):
        """Open the live SSH/DIG output in a separate window."""
        try:
            if (
                self.live_output_window is not None
                and self.live_output_window.winfo_exists()
            ):
                self.live_output_window.deiconify()
                self.live_output_window.lift()
                self.live_output_window.focus_force()
                return
            self.live_output_window = ctk.CTkToplevel(self.root)
            self.live_output_window.title(
                "INFOBLOX DIG AUTOMATION — Live Output"
            )
            self.live_output_window.geometry("1100x650")
            self.live_output_window.minsize(700, 400)
            self.live_output_window.configure(fg_color=COLOR_BG)
            self.live_output_window.grid_rowconfigure(1, weight=1)
            self.live_output_window.grid_columnconfigure(0, weight=1)
            header = ctk.CTkFrame(
                self.live_output_window,
                fg_color=COLOR_SURFACE,
                corner_radius=12
            )
            header.grid(
                row=0,
                column=0,
                sticky="ew",
                padx=14,
                pady=(14, 7)
            )
            header.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                header,
                text="⬒   LIVE OUTPUT",
                font=font(15, "bold"),
                text_color=COLOR_TEXT,
                anchor="w"
            ).grid(
                row=0,
                column=0,
                padx=16,
                pady=12,
                sticky="w"
            )
            ctk.CTkButton(
                header,
                text="Clear",
                width=90,
                height=34,
                corner_radius=10,
                fg_color=COLOR_SURFACE_ALT,
                hover_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=font(12, "bold"),
                command=self.clear_live_output
            ).grid(
                row=0,
                column=1,
                padx=12,
                pady=10
            )
            self.console = ctk.CTkTextbox(
                self.live_output_window,
                font=ctk.CTkFont(family="Consolas", size=12),
                corner_radius=12,
                fg_color="#0a0e14",
                border_width=1,
                border_color=COLOR_BORDER,
                text_color="#c9d1d9",
                wrap="none"
            )
            self.console.grid(
                row=1,
                column=0,
                sticky="nsew",
                padx=14,
                pady=(0, 14)
            )
            if self.live_output_buffer:
                self.console.insert(
                    "end",
                    self.live_output_buffer
                )
                self.console.see("end")
            self.live_output_window.protocol(
                "WM_DELETE_WINDOW",
                self.close_live_output
            )
        except Exception:
            self.live_output_window = None
            self.console = None
    def close_live_output(self):
        """Close only the live-output window; SSH/DIG execution is unaffected."""
        try:
            if (
                self.live_output_window is not None
                and self.live_output_window.winfo_exists()
            ):
                self.live_output_window.destroy()
        except Exception:
            pass
        finally:
            self.live_output_window = None
            self.console = None
    def clear_live_output(self):
        """Clear the displayed and buffered live output."""
        self.live_output_buffer = ""
        if self.console is not None:
            try:
                self.console.delete(
                    "1.0",
                    "end"
                )
            except Exception:
                pass
    # ========================================================
    # LOGGING
    # ========================================================
    def log(
        self,
        message,
        raw=False
    ):
        self.log_queue.put(
            (
                "raw" if raw else "text",
                message
            )
        )
    def process_log_queue(self):
        # Process output in small batches so a large DIG result set
        # does not block the GUI between Step 1, Step 2 and Step 3.
        processed = 0
        max_messages = 100
        try:
            while processed < max_messages:
                message_type, message = self.log_queue.get_nowait()
                if message_type == "raw":
                    display_message = message
                else:
                    display_message = "\n" + message + "\n"
                self.live_output_buffer += display_message
                if self.console is not None:
                    try:
                        self.console.insert(
                            "end",
                            display_message
                        )
                    except Exception:
                        self.console = None
                processed += 1
            if processed and self.console is not None:
                try:
                    self.console.see("end")
                except Exception:
                    self.console = None
        except queue.Empty:
            pass
        self.root.after(50, self.process_log_queue)
    # ========================================================
    # MANUAL INFOBLOX PAGER CONTROL
    # ========================================================
    def send_pager_q(self):
        """Send q + Enter to the active Infoblox SSH channel."""
        global ACTIVE_SSH_CHANNEL
        with ACTIVE_SSH_CHANNEL_LOCK:
            channel = ACTIVE_SSH_CHANNEL
        if channel is None:
            self.log(
                "No active SSH session. Start RUN ALL STEPS first."
            )
            return
        try:
            channel.send("q\n")
            self.log(
                "[MANUAL] Sent q + Enter to Infoblox pager. Continuing..."
            )
        except Exception as exc:
            self.log(
                f"[MANUAL] Failed to send q: {exc}"
            )
    # ========================================================
    # START
    # ========================================================
    def start_run(self):
        # Use the folder selected in the GUI for all generated reports.
        # Existing report functions use the module-level OUTPUT_DIR.
        global OUTPUT_DIR
        selected_output = self.output_entry.get().strip()
        if selected_output:
            OUTPUT_DIR = os.path.abspath(selected_output)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        self.open_live_output()
        self.run_button.configure(
            state="disabled"
        )
        self.clear_live_output()
        self.last_txt = None
        self.last_excel = None
        self.step1_pill.set_status("Running", COLOR_ACCENT)
        self.step2_pill.set_status("Waiting", COLOR_TEXT_MUTED)
        self.step3_pill.set_status("Waiting", COLOR_TEXT_MUTED)
        self.excel_pill.set_status("Waiting", COLOR_TEXT_MUTED)
        thread = threading.Thread(
            target=self.run_all,
            daemon=True
        )
        thread.start()
    def update_step_progress(self, step, completed, total):
        text_value = f"DIGs: {completed} / {total}"
        if step == 1:
            self.step1_progress.configure(text=text_value)
        elif step == 3:
            self.step3_progress.configure(text=text_value)
    # ========================================================
    # RUN ALL
    # ========================================================
    def run_all(self):
        session = None
        try:
            server = (
                self.server_entry.get()
                .strip()
            )
            username = (
                self.username_entry.get()
                .strip()
            )
            port = (
                self.port_entry.get()
                .strip()
            )
            authentication = (
                self.auth_var.get()
            )
            password = (
                self.password_entry.get()
            )
            key_file = (
                self.key_entry.get()
                .strip()
            )
            key_passphrase = (
                self.key_passphrase_entry.get()
            )
            # ------------------------------------------------
            # Validation
            # ------------------------------------------------
            if not server:
                raise Exception(
                    "Server / IP is required."
                )
            if not username:
                raise Exception(
                    "Username is required."
                )
            if authentication == "Password":
                if not password:
                    raise Exception(
                        "SSH password is required."
                    )
            else:
                if not key_file:
                    raise Exception(
                        "Please select a .ppk or .pem key."
                    )
                if not os.path.exists(
                    key_file
                ):
                    raise Exception(
                        "Selected key does not exist."
                    )
            # ------------------------------------------------
            # Commands
            # ------------------------------------------------
            commands = self.get_effective_commands()
            if not commands:
                raise Exception(
                    "No DIG commands found in the selected input."
                )
            self.log(
                "=" * 70
            )
            self.log(
                "INFOBLOX DIG AUTOMATION"
            )
            self.log(
                "=" * 70
            )
            self.log(
                f"Server   : {server}"
            )
            self.log(
                f"Username : {username}"
            )
            self.log(
                f"Port     : {port}"
            )
            self.log(
                f"Auth     : {authentication}"
            )
            self.log(
                f"Commands : {len(commands)}"
            )
            if authentication == "SSH Key":
                extension = os.path.splitext(
                    key_file
                )[1].upper()
                self.log(
                    f"Key Type : {extension}"
                )
            # ------------------------------------------------
            # STEP 1
            # ------------------------------------------------
            self.root.after(
                0,
                lambda total=len(commands): self.update_step_progress(
                    1, 0, total
                )
            )
            self.root.after(
                0,
                lambda total=len(commands): self.update_step_progress(
                    3, 0, 0
                )
            )
            self.log(
                "\n[STEP 1] Running Infoblox DIG..."
            )
            step1_result = execute_ssh(
                server=server,
                username=username,
                port=port,
                authentication=authentication,
                password=password,
                key_file=key_file,
                key_passphrase=key_passphrase,
                commands=commands,
                log_callback=self.log,
                keep_connection=True,
                progress_callback=lambda completed, total: self.root.after(
                    0,
                    lambda c=completed, t=total: self.update_step_progress(
                        1, c, t
                    )
                )
            )
            step1_output, session_client, session_channel, session_key_dir = step1_result
            session = (
                session_client,
                session_channel,
                session_key_dir
            )
            self.root.after(
                0,
                lambda: self.step1_pill.set_status("Completed", COLOR_SUCCESS)
            )
            # ------------------------------------------------
            # STEP 2
            # ------------------------------------------------
            self.log(
                "\n[STEP 2] Extracting unsuccessful commands..."
            )
            unsuccessful = extract_unsuccessful(
                step1_output
            )
            self.log(
                f"Unsuccessful commands: {len(unsuccessful)}"
            )
            for command in unsuccessful:
                self.log(
                    "  " + command
                )
            self.root.after(
                0,
                lambda: self.step2_pill.set_status("Completed", COLOR_SUCCESS)
            )
            # ------------------------------------------------
            # STEP 3
            # ------------------------------------------------
            self.log(
                "\n[STEP 3] Retrying unsuccessful commands..."
            )
            self.root.after(
                0,
                lambda total=len(unsuccessful): self.update_step_progress(
                    3, 0, total
                )
            )
            if unsuccessful:
                step3_output = run_step3(
                    server=server,
                    username=username,
                    port=port,
                    authentication=authentication,
                    password=password,
                    key_file=key_file,
                    key_passphrase=key_passphrase,
                    unsuccessful_commands=unsuccessful,
                    log_callback=self.log,
                    session=session,
                    progress_callback=lambda completed, total: self.root.after(
                        0,
                        lambda c=completed, t=total: self.update_step_progress(
                            3, c, t
                        )
                    )
                )
            else:
                step3_output = ""
                self.log(
                    "No commands require retry."
                )
            self.root.after(
                0,
                lambda: self.step3_pill.set_status("Completed", COLOR_SUCCESS)
            )
            # ------------------------------------------------
            # TXT
            # ------------------------------------------------
            self.log(
                "\n[OUTPUT] Creating final TXT..."
            )
            self.last_txt = create_final_file(
                server=server,
                username=username,
                authentication=authentication,
                commands=commands,
                step1_output=step1_output,
                unsuccessful=unsuccessful,
                step3_output=step3_output
            )
            self.log(
                f"TXT saved: {self.last_txt}"
            )
            # ------------------------------------------------
            # EXCEL
            # ------------------------------------------------
            self.log(
                "\n[EXCEL] Creating report..."
            )
            self.last_excel = create_excel_result(
                commands=commands,
                step1_output=step1_output,
                step3_output=step3_output,
                unsuccessful=unsuccessful,
                server=server
            )
            self.log(
                f"Excel saved: {self.last_excel}"
            )
            self.root.after(
                0,
                lambda: self.excel_pill.set_status("Created", COLOR_SUCCESS)
            )
            # ------------------------------------------------
            # COMPLETE
            # ------------------------------------------------
            self.log(
                "\n"
                + "=" * 70
            )
            self.log(
                "COMPLETED SUCCESSFULLY"
            )
            self.log(
                "=" * 70
            )
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Completed",
                    "Infoblox DIG automation completed.\n\n"
                    f"TXT:\n{self.last_txt}\n\n"
                    f"Excel:\n{self.last_excel}"
                )
            )
        except Exception as error:
            self.log(
                "\n"
                + "=" * 70
            )
            self.log(
                "ERROR"
            )
            self.log(
                "=" * 70
            )
            self.log(
                str(error)
            )
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Infoblox DIG Error",
                    str(error)
                )
            )
        finally:
            if session:
                try:
                    session[1].send("exit\n")
                except Exception:
                    pass
                try:
                    session[1].close()
                except Exception:
                    pass
                try:
                    session[0].close()
                except Exception:
                    pass
                if session[2]:
                    shutil.rmtree(
                        session[2],
                        ignore_errors=True
                    )
            global ACTIVE_SSH_CHANNEL
            with ACTIVE_SSH_CHANNEL_LOCK:
                ACTIVE_SSH_CHANNEL = None
            self.root.after(
                0,
                lambda: self.run_button.configure(
                    state="normal"
                )
            )
    # ========================================================
    # SERVER PROFILES / MULTI-INFOBLOX TEST
    # ========================================================
    def load_server_profiles(self):
        if not os.path.exists(SERVER_PROFILES_FILE):
            return []
        try:
            with open(SERVER_PROFILES_FILE, "r", encoding="utf-8") as file:
                data = json.load(file)
            if not isinstance(data, list):
                return []
            # Passwords are never kept in the server profile file.
            # Remove any password that may exist in an older profile file.
            sanitized_profiles = []
            password_found = False
            for profile in data:
                if not isinstance(profile, dict):
                    continue
                profile_copy = dict(profile)
                if "password" in profile_copy:
                    profile_copy.pop("password", None)
                    password_found = True
                sanitized_profiles.append(profile_copy)
            if password_found:
                with open(SERVER_PROFILES_FILE, "w", encoding="utf-8") as file:
                    json.dump(sanitized_profiles, file, indent=4)
            return sanitized_profiles
        except Exception as exc:
            self.log(f"[PROFILES] Unable to read profiles: {exc}")
            return []
    def save_server_profiles(self, profiles):
        # Never persist the SSH password in server_profiles.json.
        sanitized_profiles = []
        for profile in profiles:
            if not isinstance(profile, dict):
                continue
            profile_copy = dict(profile)
            profile_copy.pop("password", None)
            sanitized_profiles.append(profile_copy)
        with open(SERVER_PROFILES_FILE, "w", encoding="utf-8") as file:
            json.dump(sanitized_profiles, file, indent=4)
    def get_current_profile_data(self, name):
        return {
            "name": name,
            "server": self.server_entry.get().strip(),
            "username": self.username_entry.get().strip(),
            "port": self.port_entry.get().strip(),
            "authentication": self.auth_var.get(),
            "password": self.password_entry.get(),
            "key_file": self.key_entry.get().strip(),
            "key_passphrase": self.key_passphrase_entry.get()
        }
    def open_server_profiles(self):
        window = ctk.CTkToplevel(self.root)
        window.title("Server Profiles / Multi-Infoblox Test")
        window.geometry("980x650")
        window.minsize(860, 560)
        window.transient(self.root)
        window.configure(fg_color=COLOR_BG)
        # Header
        header = ctk.CTkFrame(window, fg_color=COLOR_SURFACE, corner_radius=14)
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(
            header,
            text="⚙   SERVER PROFILES / MULTI-INFOBLOX TEST",
            font=font(17, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=18, pady=(14, 2))
        ctk.CTkLabel(
            header,
            text="Select two or more profiles to run the same DIG command file against multiple Infoblox servers.",
            font=font(11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=18, pady=(0, 14))
        list_frame = ctk.CTkFrame(window, fg_color=COLOR_SURFACE, corner_radius=14)
        list_frame.pack(fill="both", expand=True, padx=16, pady=8)
        # Treeview styling
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Modern.Treeview",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_TEXT,
            fieldbackground=COLOR_SURFACE_ALT,
            rowheight=30,
            borderwidth=0,
            font=(FONT_FAMILY, 11)
        )
        style.configure(
            "Modern.Treeview.Heading",
            background=COLOR_BORDER,
            foreground=COLOR_TEXT,
            relief="flat",
            font=(FONT_FAMILY, 11, "bold")
        )
        style.map(
            "Modern.Treeview",
            background=[("selected", COLOR_ACCENT)],
            foreground=[("selected", "#ffffff")]
        )
        columns = ("select", "profile", "server", "username", "auth")
        tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            height=12,
            style="Modern.Treeview"
        )
        tree.heading("select", text="Select")
        tree.heading("profile", text="Profile")
        tree.heading("server", text="Server / IP")
        tree.heading("username", text="Username")
        tree.heading("auth", text="Authentication")
        tree.column("select", width=70, anchor="center")
        tree.column("profile", width=180)
        tree.column("server", width=280)
        tree.column("username", width=120)
        tree.column("auth", width=130)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=12)
        scrollbar.pack(side="right", fill="y", pady=12, padx=(0, 12))
        profiles = self.load_server_profiles()
        selected_names = set()
        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            for profile in profiles:
                tree.insert(
                    "",
                    "end",
                    iid=profile.get("name", profile.get("server", "")),
                    values=(
                        "☐",
                        profile.get("name", ""),
                        profile.get("server", ""),
                        profile.get("username", ""),
                        profile.get("authentication", "")
                    )
                )
        def toggle_selected(event=None):
            item = tree.identify_row(event.y) if event else None
            if not item:
                return
            if item in selected_names:
                selected_names.remove(item)
            else:
                selected_names.add(item)
            values = list(tree.item(item, "values"))
            values[0] = "☑" if item in selected_names else "☐"
            tree.item(item, values=values)
        tree.bind("<ButtonRelease-1>", toggle_selected)
        refresh_tree()
        button_frame = ctk.CTkFrame(window, fg_color="transparent")
        button_frame.pack(fill="x", padx=16, pady=(4, 16))
        def save_current_profile():
            dialog = ctk.CTkInputDialog(
                text="Enter a profile name:",
                title="Save Server Profile"
            )
            name = dialog.get_input()
            if not name:
                return
            name = name.strip()
            if not name:
                return
            nonlocal profiles
            profiles = [p for p in profiles if p.get("name", "") != name]
            profiles.append(self.get_current_profile_data(name))
            self.save_server_profiles(profiles)
            selected_names.add(name)
            refresh_tree()
            for item in tree.get_children():
                if item == name:
                    values = list(tree.item(item, "values"))
                    values[0] = "☑"
                    tree.item(item, values=values)
            self.log(f"[PROFILES] Saved profile: {name}")
        def load_selected_profile():
            if len(selected_names) != 1:
                messagebox.showwarning(
                    "Load Profile",
                    "Select exactly one server profile to load."
                )
                return
            name = next(iter(selected_names))
            profile = next((p for p in profiles if p.get("name") == name), None)
            if not profile:
                return
            self.server_entry.delete(0, "end")
            self.server_entry.insert(0, profile.get("server", ""))
            self.username_entry.delete(0, "end")
            self.username_entry.insert(0, profile.get("username", ""))
            self.port_entry.delete(0, "end")
            self.port_entry.insert(0, profile.get("port", "22"))
            self.auth_var.set(profile.get("authentication", "Password"))
            self.authentication_changed(self.auth_var.get())
            self.password_entry.delete(0, "end")
            self.password_entry.insert(0, profile.get("password", ""))
            self.key_entry.delete(0, "end")
            self.key_entry.insert(0, profile.get("key_file", ""))
            self.key_passphrase_entry.delete(0, "end")
            self.key_passphrase_entry.insert(0, profile.get("key_passphrase", ""))
            self.log(f"[PROFILES] Loaded profile: {name}")
        def delete_selected_profiles():
            nonlocal profiles
            if not selected_names:
                messagebox.showwarning("Delete Profile", "Select at least one profile.")
                return
            profiles = [p for p in profiles if p.get("name", "") not in selected_names]
            self.save_server_profiles(profiles)
            selected_names.clear()
            refresh_tree()
            self.log("[PROFILES] Selected profiles deleted.")
        def select_all_profiles():
            selected_names.clear()
            for item in tree.get_children():
                selected_names.add(item)
                values = list(tree.item(item, "values"))
                values[0] = "☑"
                tree.item(item, values=values)
        def clear_selection():
            selected_names.clear()
            for item in tree.get_children():
                values = list(tree.item(item, "values"))
                values[0] = "☐"
                tree.item(item, values=values)
        btn_style = dict(
            height=38,
            corner_radius=10,
            fg_color=COLOR_SURFACE_ALT,
            hover_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=font(11, "bold")
        )
        ctk.CTkButton(
            button_frame, text="SAVE CURRENT", command=save_current_profile, **btn_style
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            button_frame, text="LOAD SELECTED", command=load_selected_profile, **btn_style
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            button_frame, text="DELETE SELECTED", command=delete_selected_profiles, **btn_style
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            button_frame, text="SELECT ALL", command=select_all_profiles, **btn_style
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            button_frame, text="CLEAR", command=clear_selection, **btn_style
        ).pack(side="left", padx=4)
        def run_selected():
            selected_profiles = [
                p for p in profiles
                if p.get("name", "") in selected_names
            ]
            if not selected_profiles:
                messagebox.showwarning("Multi-Infoblox Test", "Select at least one server profile.")
                return

            # Passwords are intentionally not stored in server_profiles.json.
            # For Multi-Infoblox Test, ask once for a common SSH password when
            # any selected server uses Password authentication. The password
            # is kept only in memory for this run and is never written to disk.
            password_profiles = [
                p for p in selected_profiles
                if p.get("authentication", "Password") == "Password"
            ]
            if password_profiles:
                common_password = simpledialog.askstring(
                    "Multi-Infoblox Test",
                    "Enter the common SSH password for the selected Password-authentication servers:",
                    parent=window,
                    show="*"
                )
                if common_password is None:
                    return
                if not common_password:
                    messagebox.showwarning(
                        "Multi-Infoblox Test",
                        "SSH password is required for the selected Password-authentication servers."
                    )
                    return

                # Runtime-only password assignment. This does not modify the
                # profile file because save_server_profiles() is not called.
                for profile in password_profiles:
                    profile["password"] = common_password

            try:
                commands = self.get_effective_commands()
            except Exception as exc:
                messagebox.showerror("Multi-Infoblox Test", str(exc))
                return
            if not commands:
                messagebox.showwarning("Multi-Infoblox Test", "No DIG commands found.")
                return
            self.open_live_output()
            window.destroy()
            self.run_multi_server(selected_profiles, commands)
        ctk.CTkButton(
            button_frame,
            text="▶   RUN SELECTED SERVERS",
            height=40,
            corner_radius=10,
            fg_color=COLOR_SUCCESS,
            hover_color="#059669",
            text_color="#ffffff",
            font=font(12, "bold"),
            command=run_selected
        ).pack(side="right", padx=4)
    def run_multi_server(self, profiles, commands):
        """Run the existing Step 1 -> Step 2 -> Step 3 flow in parallel per profile."""
        # Use the folder selected in the GUI for all generated reports,
        # including the multi-server command-comparison report.
        global OUTPUT_DIR
        selected_output = self.output_entry.get().strip()
        if selected_output:
            OUTPUT_DIR = os.path.abspath(selected_output)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        self.run_button.configure(state="disabled")
        self.profiles_button.configure(state="disabled")

        # ------------------------------------------------------------
        # Existing summary window
        # ------------------------------------------------------------
        summary_window = ctk.CTkToplevel(self.root)
        summary_window.title("Multi-Infoblox Test Results")
        summary_window.geometry("960x560")
        summary_window.minsize(820, 480)
        summary_window.transient(self.root)
        summary_window.configure(fg_color=COLOR_BG)

        header = ctk.CTkFrame(
            summary_window,
            fg_color=COLOR_SURFACE,
            corner_radius=14
        )
        header.pack(fill="x", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            header,
            text="MULTI-INFOBLOX TEST RESULTS",
            font=font(17, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=18, pady=(14, 2))

        ctk.CTkLabel(
            header,
            text=f"Commands: {len(commands)}    Servers: {len(profiles)}",
            font=font(11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=18, pady=(0, 14))

        table_frame = ctk.CTkFrame(
            summary_window,
            fg_color=COLOR_SURFACE,
            corner_radius=14
        )
        table_frame.pack(
            fill="both",
            expand=True,
            padx=16,
            pady=8
        )

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Multi.Treeview",
            background=COLOR_SURFACE_ALT,
            foreground=COLOR_TEXT,
            fieldbackground=COLOR_SURFACE_ALT,
            rowheight=30,
            borderwidth=0,
            font=(FONT_FAMILY, 11)
        )
        style.configure(
            "Multi.Treeview.Heading",
            background=COLOR_BORDER,
            foreground=COLOR_TEXT,
            relief="flat",
            font=(FONT_FAMILY, 11, "bold")
        )

        columns = (
            "server",
            "success",
            "failed",
            "timeout",
            "total"
        )

        result_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            style="Multi.Treeview"
        )

        headings = {
            "server": "Server",
            "success": "Success",
            "failed": "Failed",
            "timeout": "Timeout",
            "total": "Total"
        }

        widths = {
            "server": 420,
            "success": 110,
            "failed": 110,
            "timeout": 110,
            "total": 110
        }

        for col in columns:
            result_tree.heading(col, text=headings[col])
            result_tree.column(
                col,
                width=widths[col],
                anchor="center" if col != "server" else "w"
            )

        result_tree.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=12
        )

        status_label = ctk.CTkLabel(
            summary_window,
            text="Running in parallel...",
            font=font(12, "bold"),
            text_color=COLOR_ACCENT
        )
        status_label.pack(pady=(0, 14))

        # ------------------------------------------------------------
        # NEW: separate live dashboard for parallel server execution
        # ------------------------------------------------------------
        dashboard_window = ctk.CTkToplevel(self.root)
        dashboard_window.title("Multi-Infoblox Test — Live Dashboard")
        dashboard_window.geometry("1200x760")
        dashboard_window.minsize(900, 520)
        dashboard_window.configure(fg_color=COLOR_BG)
        dashboard_window.grid_rowconfigure(1, weight=1)
        dashboard_window.grid_columnconfigure(0, weight=1)

        dashboard_header = ctk.CTkFrame(
            dashboard_window,
            fg_color=COLOR_SURFACE,
            corner_radius=14
        )
        dashboard_header.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=14,
            pady=(14, 8)
        )

        ctk.CTkLabel(
            dashboard_header,
            text="MULTI-INFOBLOX LIVE EXECUTION",
            font=font(17, "bold"),
            text_color=COLOR_TEXT
        ).pack(anchor="w", padx=18, pady=(14, 2))

        ctk.CTkLabel(
            dashboard_header,
            text=(
                f"Running {len(profiles)} servers simultaneously  •  "
                f"Total DIG commands per server: {len(commands)}"
            ),
            font=font(11),
            text_color=COLOR_TEXT_MUTED
        ).pack(anchor="w", padx=18, pady=(0, 14))

        dashboard_scroll = ctk.CTkScrollableFrame(
            dashboard_window,
            fg_color="transparent",
            corner_radius=0
        )
        dashboard_scroll.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=14,
            pady=(0, 14)
        )
        dashboard_scroll.grid_columnconfigure(0, weight=1)
        dashboard_scroll.grid_columnconfigure(1, weight=1)

        dashboard_widgets = {}

        for index, profile in enumerate(profiles):
            server_name = profile.get("server", "") or profile.get("name", "") or f"Server {index + 1}"

            card = ctk.CTkFrame(
                dashboard_scroll,
                fg_color=COLOR_SURFACE,
                corner_radius=14,
                border_width=1,
                border_color=COLOR_BORDER
            )
            card.grid(
                row=index // 2,
                column=index % 2,
                sticky="nsew",
                padx=6,
                pady=6
            )
            card.grid_columnconfigure(0, weight=1)
            card.grid_rowconfigure(3, weight=1)

            ctk.CTkLabel(
                card,
                text=f"🖥  {server_name}",
                font=font(13, "bold"),
                text_color=COLOR_TEXT,
                anchor="w"
            ).grid(
                row=0,
                column=0,
                padx=14,
                pady=(12, 2),
                sticky="ew"
            )

            server_status = ctk.CTkLabel(
                card,
                text="Waiting...",
                font=font(11, "bold"),
                text_color=COLOR_ACCENT,
                anchor="w"
            )
            server_status.grid(
                row=1,
                column=0,
                padx=14,
                pady=(0, 4),
                sticky="ew"
            )

            progress_text = ctk.CTkLabel(
                card,
                text=f"Commands executed: 0 / {len(commands)}",
                font=font(10),
                text_color=COLOR_TEXT_MUTED,
                anchor="w"
            )
            progress_text.grid(
                row=2,
                column=0,
                padx=14,
                pady=(0, 4),
                sticky="ew"
            )

            progress_bar = ctk.CTkProgressBar(
                card,
                height=10,
                corner_radius=6,
                fg_color=COLOR_SURFACE_ALT,
                progress_color=COLOR_ACCENT
            )
            progress_bar.set(0)
            progress_bar.grid(
                row=3,
                column=0,
                padx=14,
                pady=(2, 8),
                sticky="ew"
            )

            output_box = ctk.CTkTextbox(
                card,
                height=230,
                corner_radius=10,
                fg_color="#080b0f",
                border_width=1,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=ctk.CTkFont(family="Consolas", size=10),
                wrap="none"
            )
            output_box.grid(
                row=4,
                column=0,
                padx=14,
                pady=(0, 14),
                sticky="nsew"
            )

            dashboard_widgets[index] = {
                "server": server_name,
                "status": server_status,
                "progress_text": progress_text,
                "progress_bar": progress_bar,
                "output": output_box
            }

            output_box.insert(
                "end",
                f"[{server_name}] Waiting to start...\n"
            )

        def dashboard_alive():
            try:
                return (
                    dashboard_window is not None
                    and dashboard_window.winfo_exists()
                )
            except Exception:
                return False

        def append_dashboard_output(server_index, message):
            """Append one server's live SSH output to its dashboard box."""
            def update():
                try:
                    if not dashboard_alive():
                        return
                    widget = dashboard_widgets[server_index]["output"]
                    widget.insert("end", message)
                    widget.see("end")
                except Exception:
                    pass

            try:
                self.root.after(0, update)
            except Exception:
                pass

        def update_dashboard_status(
            server_index,
            text_value,
            color=COLOR_ACCENT
        ):
            def update():
                try:
                    if not dashboard_alive():
                        return
                    dashboard_widgets[server_index]["status"].configure(
                        text=text_value,
                        text_color=color
                    )
                except Exception:
                    pass

            try:
                self.root.after(0, update)
            except Exception:
                pass

        def update_dashboard_progress(
            server_index,
            completed,
            total,
            phase="Step 1"
        ):
            def update():
                try:
                    if not dashboard_alive():
                        return
                    widget_set = dashboard_widgets[server_index]
                    ratio = (
                        min(1.0, max(0.0, completed / total))
                        if total
                        else 0
                    )
                    widget_set["progress_bar"].set(ratio)
                    widget_set["progress_text"].configure(
                        text=(
                            f"{phase} commands executed: "
                            f"{completed} / {total}"
                        )
                    )
                except Exception:
                    pass

            try:
                self.root.after(0, update)
            except Exception:
                pass

        def server_log_callback(server_index, message, raw=False):
            # Preserve the existing central Live Output behavior.
            self.log(message, raw=raw)

            # Add the same live output to the server-specific dashboard.
            append_dashboard_output(server_index, message)

        def add_row(server, success, failed, timeout, total):
            self.root.after(
                0,
                lambda: result_tree.insert(
                    "",
                    "end",
                    values=(server, success, failed, timeout, total)
                )
            )

        comparison_results = {}
        comparison_lock = threading.Lock()
        completed_lock = threading.Lock()
        report_lock = threading.Lock()
        completed_servers = 0

        def store_comparison(server, command, status):
            with comparison_lock:
                comparison_results.setdefault(
                    normalize_command(command),
                    {}
                )[server] = status

        def make_server_reports(
            server,
            username,
            authentication,
            commands_for_report,
            step1_output,
            unsuccessful,
            step3_output
        ):
            """
            Keep report generation serialized so the existing timestamp-based
            report filenames do not collide while SSH execution runs in parallel.
            """
            with report_lock:
                try:
                    txt_report = create_final_file(
                        server=server,
                        username=username,
                        authentication=authentication,
                        commands=commands_for_report,
                        step1_output=step1_output,
                        unsuccessful=unsuccessful,
                        step3_output=step3_output
                    )

                    excel_report = create_excel_result(
                        commands=commands_for_report,
                        step1_output=step1_output,
                        step3_output=step3_output,
                        unsuccessful=unsuccessful,
                        server=server
                    )

                    # The existing report functions use timestamp-only names.
                    # Move each generated file to a server-specific name
                    # immediately so parallel servers never overwrite a report.
                    safe_server = re.sub(
                        r"[^A-Za-z0-9._-]+",
                        "_",
                        server
                    ).strip("_") or "server"

                    def unique_server_path(path):
                        base, ext = os.path.splitext(path)
                        candidate = f"{base}_{safe_server}{ext}"
                        counter = 2
                        while os.path.exists(candidate):
                            candidate = f"{base}_{safe_server}_{counter}{ext}"
                            counter += 1
                        return candidate

                    txt_target = unique_server_path(txt_report)
                    excel_target = unique_server_path(excel_report)

                    if os.path.abspath(txt_report) != os.path.abspath(txt_target):
                        os.replace(txt_report, txt_target)
                        txt_report = txt_target

                    if os.path.abspath(excel_report) != os.path.abspath(excel_target):
                        os.replace(excel_report, excel_target)
                        excel_report = excel_target

                    self.log(
                        f"[MULTI TEST] TXT saved: {txt_report}"
                    )
                    self.log(
                        f"[MULTI TEST] Excel saved: {excel_report}"
                    )

                except Exception as report_error:
                    self.log(
                        f"[MULTI TEST] REPORT ERROR for {server}: "
                        f"{report_error}"
                    )

        def server_worker(server_index, profile):
            nonlocal completed_servers

            server = profile.get("server", "")
            username = profile.get("username", "")
            port = profile.get("port", "22")
            authentication = profile.get("authentication", "Password")
            password = profile.get("password", "")
            key_file = profile.get("key_file", "")
            key_passphrase = profile.get("key_passphrase", "")

            session = None

            update_dashboard_status(
                server_index,
                "Connecting...",
                COLOR_WARNING
            )
            append_dashboard_output(
                server_index,
                "\n" + "=" * 60 + "\n"
            )
            append_dashboard_output(
                server_index,
                f"[MULTI TEST] {server}\n"
            )
            append_dashboard_output(
                server_index,
                "=" * 60 + "\n"
            )

            self.log("\n" + "=" * 70)
            self.log(f"[MULTI TEST] {server}")
            self.log("=" * 70)

            try:
                update_dashboard_status(
                    server_index,
                    "Running Step 1...",
                    COLOR_ACCENT
                )

                def step1_progress(done, total):
                    update_dashboard_progress(
                        server_index,
                        done,
                        total,
                        "Step 1"
                    )
                    update_dashboard_status(
                        server_index,
                        f"Running Step 1... {done}/{total}",
                        COLOR_ACCENT
                    )

                step1 = execute_ssh(
                    server=server,
                    username=username,
                    port=port,
                    authentication=authentication,
                    password=password,
                    key_file=key_file,
                    key_passphrase=key_passphrase,
                    commands=commands,
                    log_callback=lambda message, raw=False, idx=server_index: (
                        server_log_callback(idx, message, raw=raw)
                    ),
                    keep_connection=True,
                    progress_callback=step1_progress
                )

                if not step1:
                    for command in commands:
                        store_comparison(
                            server,
                            command,
                            "CONNECTION TIMED OUT"
                        )

                    update_dashboard_progress(
                        server_index,
                        len(commands),
                        len(commands),
                        "Step 1"
                    )
                    update_dashboard_status(
                        server_index,
                        f"Connection timed out — {len(commands)}/{len(commands)}",
                        COLOR_DANGER
                    )
                    add_row(
                        server,
                        0,
                        0,
                        len(commands),
                        len(commands)
                    )
                    return

                step1_output, client, channel, key_dir = step1
                session = (client, channel, key_dir)

                unsuccessful = extract_unsuccessful(step1_output)
                step3_output = ""

                update_dashboard_status(
                    server_index,
                    (
                        f"Step 1 complete — "
                        f"{len(commands)}/{len(commands)}"
                    ),
                    COLOR_SUCCESS
                )
                update_dashboard_progress(
                    server_index,
                    len(commands),
                    len(commands),
                    "Step 1"
                )

                if unsuccessful:
                    retry_total = len(unsuccessful)
                    update_dashboard_status(
                        server_index,
                        f"Retrying unsuccessful commands: 0/{retry_total}",
                        COLOR_WARNING
                    )

                    def step3_progress(done, total):
                        update_dashboard_status(
                            server_index,
                            f"Retrying unsuccessful commands: {done}/{total}",
                            COLOR_WARNING
                        )
                        append_dashboard_output(
                            server_index,
                            f"\n[STEP 3] Retry progress: {done}/{total}\n"
                        )

                    step3_output = run_step3(
                        server=server,
                        username=username,
                        port=port,
                        authentication=authentication,
                        password=password,
                        key_file=key_file,
                        key_passphrase=key_passphrase,
                        unsuccessful_commands=unsuccessful,
                        log_callback=lambda message, raw=False, idx=server_index: (
                            server_log_callback(idx, message, raw=raw)
                        ),
                        session=session,
                        progress_callback=step3_progress
                    )

                # --------------------------------------------------------
                # Existing TXT + Excel report logic.
                # It is serialized only to protect the existing
                # timestamp-based filenames.
                # --------------------------------------------------------
                make_server_reports(
                    server=server,
                    username=username,
                    authentication=authentication,
                    commands_for_report=commands,
                    step1_output=step1_output,
                    unsuccessful=unsuccessful,
                    step3_output=step3_output
                )

                step1_results = parse_actual_results(step1_output)
                step3_statuses = parse_step3_statuses(
                    step3_output,
                    unsuccessful
                )

                step3_by_key = {}
                for i, command in enumerate(unsuccessful):
                    if i < len(step3_statuses):
                        step3_by_key[command_key(command)] = step3_statuses[i]

                success = 0
                failed = 0
                timeout = 0

                for command in commands:
                    key = command_key(command)

                    if key in step3_by_key:
                        result_status = step3_by_key[key]
                    else:
                        item = step1_results.get(key)
                        result_status = (
                            item["status"]
                            if item
                            else "NO RESULT"
                        )

                    if result_status == "CONNECTION TIMED OUT":
                        timeout += 1
                    elif result_status in ("NOERROR", "NXDOMAIN"):
                        success += 1
                    else:
                        failed += 1

                    store_comparison(
                        server,
                        command,
                        result_status
                    )

                add_row(
                    server,
                    success,
                    failed,
                    timeout,
                    len(commands)
                )

                update_dashboard_progress(
                    server_index,
                    len(commands),
                    len(commands),
                    "Step 1"
                )
                update_dashboard_status(
                    server_index,
                    (
                        f"Completed — "
                        f"{len(commands)}/{len(commands)} commands  •  "
                        f"Success: {success}  •  "
                        f"Failed: {failed}  •  "
                        f"Timeout: {timeout}"
                    ),
                    COLOR_SUCCESS
                )

            except Exception as exc:
                self.log(
                    f"[MULTI TEST] {server} ERROR: {exc}"
                )

                append_dashboard_output(
                    server_index,
                    f"\n[MULTI TEST ERROR] {exc}\n"
                )

                for command in commands:
                    store_comparison(
                        server,
                        command,
                        "ERROR"
                    )

                add_row(
                    server,
                    0,
                    len(commands),
                    0,
                    len(commands)
                )

                update_dashboard_status(
                    server_index,
                    "ERROR",
                    COLOR_DANGER
                )

            finally:
                if session:
                    try:
                        session[1].send("exit\n")
                    except Exception:
                        pass

                    try:
                        session[1].close()
                    except Exception:
                        pass

                    try:
                        session[0].close()
                    except Exception:
                        pass

                    if session[2]:
                        shutil.rmtree(
                            session[2],
                            ignore_errors=True
                        )

                global ACTIVE_SSH_CHANNEL
                with ACTIVE_SSH_CHANNEL_LOCK:
                    ACTIVE_SSH_CHANNEL = None

                with completed_lock:
                    completed_servers += 1
                    current_completed = completed_servers

                self.root.after(
                    0,
                    lambda c=current_completed, total=len(profiles): (
                        status_label.configure(
                            text=(
                                f"Completed servers: {c} / {total}"
                            )
                        )
                    )
                )

        def create_command_comparison():
            try:
                timestamp = datetime.datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )

                comparison_file = os.path.join(
                    OUTPUT_DIR,
                    f"infoblox_multi_test_command_comparison_{timestamp}.xlsx"
                )

                wb = Workbook()
                ws = wb.active
                ws.title = "Command Comparison"

                server_names = [
                    profile.get("server", "")
                    for profile in profiles
                ]

                headers = (
                    ["DIG Command"]
                    + server_names
                    + ["Result"]
                )

                ws.append(headers)

                for cell in ws[1]:
                    cell.font = Font(
                        bold=True,
                        color="FFFFFF"
                    )
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="1F4E78"
                    )
                    cell.alignment = Alignment(
                        horizontal="center"
                    )

                for command in commands:
                    clean_command = normalize_command(command)
                    server_results = comparison_results.get(
                        clean_command,
                        {}
                    )

                    results_for_command = [
                        server_results.get(
                            server,
                            "NO RESULT"
                        )
                        for server in server_names
                    ]

                    same_on_all = (
                        len(set(results_for_command)) == 1
                        and results_for_command[0] != "NO RESULT"
                    )

                    comparison_status = (
                        "✅ Same on all"
                        if same_on_all
                        else "❌ Not Same on all"
                    )

                    ws.append(
                        [clean_command]
                        + results_for_command
                        + [comparison_status]
                    )

                ws.column_dimensions["A"].width = 55

                for column_index in range(
                    2,
                    len(server_names) + 2
                ):
                    ws.column_dimensions[
                        get_column_letter(column_index)
                    ].width = 30

                result_column_index = len(server_names) + 2

                ws.column_dimensions[
                    get_column_letter(result_column_index)
                ].width = 22

                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                wb.save(comparison_file)

                self.log(
                    "[MULTI TEST] Command Comparison Excel saved: "
                    f"{comparison_file}"
                )

                if os.name == "nt":
                    os.startfile(comparison_file)
                elif platform.system() == "Darwin":
                    subprocess.Popen(
                        ["open", comparison_file]
                    )
                else:
                    subprocess.Popen(
                        ["xdg-open", comparison_file]
                    )

                status_label.configure(
                    text=(
                        f"Completed: {completed_servers} / "
                        f"{len(profiles)} servers - "
                        "Command Comparison created"
                    )
                )

            except Exception as comparison_error:
                self.log(
                    "[MULTI TEST] COMMAND COMPARISON ERROR: "
                    f"{comparison_error}"
                )

        def run_all_servers_parallel():
            """
            Start one independent worker thread per selected server.
            Each server has its own SSH session and live dashboard output.
            """
            threads = []

            for index, profile in enumerate(profiles):
                thread = threading.Thread(
                    target=server_worker,
                    args=(index, profile),
                    daemon=True
                )
                threads.append(thread)
                thread.start()

            # Wait for all server workers without blocking Tk's main thread.
            for thread in threads:
                thread.join()

            self.root.after(
                0,
                create_command_comparison
            )

            self.root.after(
                0,
                lambda: self.run_button.configure(
                    state="normal"
                )
            )

            self.root.after(
                0,
                lambda: self.profiles_button.configure(
                    state="normal"
                )
            )

        # Start all selected servers simultaneously.
        threading.Thread(
            target=run_all_servers_parallel,
            daemon=True
        ).start()

    # ========================================================
    # OPEN OUTPUT
    # ========================================================
    def open_output(self):
        folder = (
            self.output_entry.get()
            .strip()
            or OUTPUT_DIR
        )
        if not os.path.exists(
            folder
        ):
            os.makedirs(
                folder,
                exist_ok=True
            )
        if os.name == "nt":
            os.startfile(
                folder
            )
        elif platform.system() == "Darwin":
            subprocess.Popen(
                ["open", folder]
            )
        else:
            subprocess.Popen(
                ["xdg-open", folder]
            )
# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    # Run the quarterly license check before creating the main GUI.
    # All startup exceptions are written to a local log as well, so a
    # packaged EXE does not fail silently if the GUI itself raises an error.
    try:
        if not ensure_quarterly_license():
            raise SystemExit(0)

        root = ctk.CTk()
        app = InfobloxDigApp(root)
        root.mainloop()

    except SystemExit:
        raise
    except Exception as exc:
        error_text = traceback.format_exc()
        try:
            log_file = os.path.join(BASE_DIR, "startup_error.log")
            with open(log_file, "w", encoding="utf-8") as f:
                f.write(error_text)
        except Exception:
            pass

        try:
            error_root = tk.Tk()
            error_root.withdraw()
            messagebox.showerror(
                "Application Startup Error",
                f"The application could not start.\n\n{exc}\n\n"
                "A detailed log was written to startup_error.log.",
                parent=error_root,
            )
            error_root.destroy()
        except Exception:
            print(error_text)
