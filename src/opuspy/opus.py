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

from contextlib import contextmanager

import pythoncom


def say_hello_from_opuspy():
    logger.info("Hello from opuspy")


@contextmanager
def sap_connection(timeout=30, interval=1):
    """
    Attempt to bind to an existing SAP GUI session within a given timeout.
    Yields the session once available. Optionally, place any cleanup code
    after the yield.
    """
    start_time = time.time()
    session = None

    while True:
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            scripting_engine = sap_gui.GetScriptingEngine
            if scripting_engine is not None:
                connection = scripting_engine.Connections(0)
                session = connection.Sessions(0)
                break
        except pythoncom.com_error:
            pass

        if (time.time() - start_time) > timeout:
            raise RuntimeError("Timed out waiting for SAP GUI readiness.")

        time.sleep(interval)

    # Now that session is found, yield control
    yield session


def start_opus(pam_path, user, sapshcut_path):
    """
    Start SAP session with SAPSHCUT.exe
    :param pam_path: str, path to PAM file
    :param user: str, user to log in as
    :param sapshcut_path: str, path to SAPSHCUT.exe
    """

    # Unpack credentials
    username, password = get_credentials(pam_path, user, fagsystem="opus")

    if not username or not password:
        logger.error("Failed to retrieve credentials.")
        return None

    command_args = [
        str(sapshcut_path),
        "-system=P02",
        "-client=400",
        f"-user={username}",
        f"-pw={password}",
    ]

    subprocess.run(command_args, check=False)  # noqa: S603

    try:
        with sap_connection() as session:
            # Check if SAP with ID /app/con[0]/ses[0]/wnd[1]/usr is open
            element_id = "/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT"
            try:
                element = session.findById(element_id)
                if element.text == "Nyt password":
                    logger.info("Detected password reset prompt in SAP.")

                    backup_old_password(pam_path=pam_path, user=user)
                    new_password = generate_new_password(17)
                    save_new_password(
                        new_password=new_password,
                        pam_path=pam_path,
                        user=user,
                        fagsystem="opus",
                    )

                    # Write new password to SAP
                    session.findById(
                        "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCODE"
                    ).text = new_password
                    session.findById(
                        "/app/con[0]/ses[0]/wnd[1]/usr/pwdRSYST-NCOD2"
                    ).text = new_password

                    # Press OK
                    session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").press()
                    logger.info("Password updated successfully in SAP.")
                    time.sleep(1)

                else:
                    logger.info("Password reset prompt not detected.")

            except Exception as e:
                logger.error(f"Error interacting with SAP element {element_id}: {e}")
                raise

        return session

    except Exception as e:
        logger.error(f"Failed to start SAP session: {e}")
        return None


# Example usage
if __name__ == "__main__":
    start_opus()
