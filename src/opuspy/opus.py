import subprocess
import time

import win32com.client  # pywin32

from brkrpautils import (
    get_credentials,
    backup_old_password,
    generate_new_password,
    save_new_password,
)


def say_hello_from_opuspy():
    print("Hello from opuspy")


def start_opus(pam_path, user, sapshcut_path):
    """
    Start SAP session with SAPSHCUT.exe
    :param pam_path: str, path to PAM file
    :param user: str, user to log in as
    :param sapshcut_path: str, path to SAPSHCUT.exe
    """

    # unpacking
    username, password = get_credentials(pam_path, user, fagsystem="opus")

    if not username or not password:
        return None

    command_args = [
        str(sapshcut_path),
        "-system=P02",
        "-client=400",
        f"-user={username}",
        f"-pw={password}",
    ]

    subprocess.run(command_args, check=False)  # noqa: S603
    time.sleep(1)

    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not isinstance(SapGuiAuto, win32com.client.CDispatch):
            return

        application = SapGuiAuto.GetScriptingEngine
        if not isinstance(application, win32com.client.CDispatch):
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not isinstance(connection, win32com.client.CDispatch):
            application = None
            SapGuiAuto = None
            return

        if connection.DisabledByServer:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not isinstance(session, win32com.client.CDispatch):
            connection = None
            application = None
            SapGuiAuto = None
            return

        if session.Busy:
            session = None
            connection = None
            application = None
            return

        if session.Info.IsLowSpeedConnection:
            session = None
            connection = None
            application = None
            return

        # ---------------------------------------------------------------------------- #
        #                                  Script code                                 #
        # ---------------------------------------------------------------------------- #

        # Check if SAP with ID /app/con[0]/ses[0]/wnd[1]/usr is open
        if (
            session.findById("/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT").text
            == "Nyt password"
        ):
            backup_old_password(pam_path=pam_path, user=user)
            new_password = generate_new_password(17)
            save_new_password(
                new_password=new_password,
                pam_path=pam_path,
                user=user,
                fagsystem="opus",
            )

            # write new password to SAP
            session.findById(
                "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCODE"
            ).text = new_password
            # repeat new password
            session.findById(
                "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCOD2"
            ).text = new_password
            # press OK
            session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(1)

        return session, connection, application, SapGuiAuto

    except Exception as e:
        print(f"Error: {e}")
        return None
