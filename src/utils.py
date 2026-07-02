import logging
import os
import time

from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class RelativePathFormatter(logging.Formatter):
    def __init__(self, *args, project_root=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.project_root = str(project_root) if project_root else os.getcwd()

    def format(self, record):
        if record.pathname:
            try:
                record.rel_pathname = os.path.relpath(record.pathname, self.project_root)
            except Exception:
                record.rel_pathname = record.pathname
        else:
            record.rel_pathname = "unknown"
        return super().format(record)


def scroll_and_click(driver: WebDriver, wait: WebDriverWait[WebDriver], locator: tuple[str, str], scroll: bool = True, to_be_clickable: bool = True, js_click: bool = False) -> None:
    try:
        btn = wait.until(EC.presence_of_element_located(locator))
        if scroll:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)  # type: ignore[reportUnknownMemberType]

        if to_be_clickable:
            btn = wait.until(EC.element_to_be_clickable(locator))

        if js_click:
            driver.execute_script("arguments[0].click();", btn)  # type: ignore[reportUnknownMemberType]
        else:
            btn.click()
    except Exception as e:
        print(f"Click failed/intercepted ({e}), waiting 2 seconds and retrying...")
        time.sleep(2)
        try:
            btn = wait.until(EC.presence_of_element_located(locator))
            if scroll:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)  # type: ignore[reportUnknownMemberType]

            if to_be_clickable:
                btn = wait.until(EC.element_to_be_clickable(locator))

            if js_click:
                driver.execute_script("arguments[0].click();", btn)  # type: ignore[reportUnknownMemberType]
            else:
                btn.click()
        except Exception as e2:
            print(f"Retry failed ({e2}), using JavaScript click fallback...")
            try:
                btn = wait.until(EC.presence_of_element_located(locator))
                driver.execute_script("arguments[0].click();", btn)  # type: ignore[reportUnknownMemberType]
            except Exception as e3:
                print(f"JavaScript fallback failed: {e3}")
                raise e3