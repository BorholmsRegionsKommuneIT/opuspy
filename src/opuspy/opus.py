import subprocess
import time
import winreg

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

    with sap_connection() as session:
        # Check if SAP with ID /app/con[0]/ses[0]/wnd[1]/usr is open to determine if password reset prompt is present
        element_id = "/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT"
        try:
            element = session.findById(element_id)
        except Exception as e:
            return  # password prompt not present, continuing as normal

        if element.text == "Nyt password":
            try:
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
            except Exception as e:
                logger.error(f"Error while trying to change password {element_id}: {e}")
                raise e

        else:
            logger.info("element_id found, but text did not match 'Nyt password'.")
            raise RuntimeError(
                "element_id found, but text did not match 'Nyt password'."
            )

def is_sap_scripting_allowed() -> bool:
    """
    TRUE if scripting is allowed, FALSE if not.
    """
    try:
        # Define the registry path
        registry_path = (
            r"SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Security"
        )
        key_name = "UserScripting"

        # Open the registry key
        reg_key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_READ
        )

        # Get the value of the UserScripting key
        value, regtype = winreg.QueryValueEx(reg_key, key_name)

        # if value is 1, scripting is allowed
        if value == 1:
            return True
        else:
            return False

        # Close the registry key
        winreg.CloseKey(reg_key)
    except FileNotFoundError:
        print("The specified registry key or value does not exist.")
    except PermissionError:
        print("Permission denied. Please run the script as an administrator.")
    except Exception as e:
        print(f"An error occurred: {e}")



if __name__ == "__main__":
    if is_sap_scripting_allowed():
        start_opus()
    else:
        print("SAP scripting is not allowed. Please ask admin to enable scripting in registry")
