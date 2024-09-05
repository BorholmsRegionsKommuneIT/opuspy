import subprocess
import time

from loguru import logger
import win32com.client  # pywin32

from brkrpautils import (
    get_credentials,
    backup_old_password,
    generate_new_password,
    save_new_password,
)


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
        logger.error("Failed to retrieve credentials for robot", exc_info=True)
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
        sap = win32com.client.GetObject("SAPGUI")
        app = sap.GetScriptingEngine
        connection = app.Connections(0)
        session = connection.sessions(0)

        # Check if SAP with ID /app/con[0]/ses[0]/wnd[1]/usr is open
        if (
            session.findById("/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT").text
            == "Nyt password"
        ):
            logger.info("Password change required")
            backup_old_password(pam_path=pam_path, user=user)
            logger.info("Backup of old password saved")
            new_password = generate_new_password(17)
            save_new_password(
                new_password=new_password,
                pam_path=pam_path,
                user=user,
                fagsystem="opus",
            )
            logger.info(f"New password saved: {new_password[:3]}{'***'}")

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

        return session

    except Exception:
        logger.error("Failed to start SAP session", exc_info=True)
        return None
