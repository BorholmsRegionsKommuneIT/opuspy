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


@contextmanager
def sap_connection():
    """
    Context manager for handling the SAP GUI scripting connection.
    Initializes the SAP session and ensures proper cleanup.
    """
    session = None
    try:
        # Create the SAP GUI scripting connection
        logger.info("Starting SAP GUI connection.")
        sap_gui = win32com.client.Dispatch("SapGui.ScriptingCtrl")
        application = sap_gui.GetScriptingEngine
        connection = application.OpenConnection(
            "SAP Connection Name", True
        )  # Replace with your SAP connection name
        session = connection.Children(0)
        yield session
    except Exception as e:
        logger.error("Error while connecting to SAP: {}", e)
        raise
    finally:
        # Properly close the session or connection if applicable
        if session:
            logger.info("Closing SAP session.")
            session = None


def start_opus():
    """
    Starts the Opus session using the SAP connection and checks for the existence of an element.
    """
    with sap_connection() as session:
        try:
            # Attempt to find the element and check its text
            element = session.findById(
                "/app/con[0]/ses[0]/wnd[1]/usr/lblRSYST-NCODE_TEXT"
            )
            if (
                element.text == "EXPECTED_VALUE"
            ):  # Replace with the actual value you are comparing
                logger.info("Element exists and matches expected value.")
            else:
                logger.warning("Element exists but does not match the expected value.")
        except Exception as e:
            # Handle the case where the element does not exist or another error occurs
            logger.error("Element does not exist or another error occurred.")
            logger.debug(f"Exception details: {e}")


# Example usage
if __name__ == "__main__":
    start_opus()
