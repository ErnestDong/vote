#%%
import logging
from pathlib import Path
from time import sleep

import pandas as pd
import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.proxy import Proxy, ProxyType
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from config import ANALYSTS, GROUP, PROXIES

logging.basicConfig(level=logging.INFO)
sleep_magic = 5


class FVote:
    def __init__(self, user: str, password: str, groupname: str, option=None) -> None:
        self.user = user
        self.password = password
        if option is None:
            option = selenium.webdriver.ChromeOptions()
        self.driver = selenium.webdriver.Chrome(options=option)
        self.group = groupname

    def login(self):
        # NOTE: 如果不在前台开着应用会崩溃
        user = self.user
        password = self.password
        driver = self.driver
        driver.get("https://voting.institutionalinvestor.com/welcome")
        WebDriverWait(driver, sleep_magic).until(
            EC.presence_of_element_located((By.ID, "btnLogin"))
        ).click()
        sleep(sleep_magic)
        WebDriverWait(driver, sleep_magic).until(
            EC.presence_of_element_located((By.TAG_NAME, "input"))
        )
        driver.find_elements(By.TAG_NAME, "input")[0].send_keys(user)
        driver.find_elements(By.TAG_NAME, "input")[1].send_keys(password)
        driver.find_element(By.TAG_NAME, "button").click()
        logging.info(f"Login successfully as {user}")
        # "VOTE"
        sleep(sleep_magic * 2)
        driver.find_elements(By.TAG_NAME, "button")[2].click()
        self.driver = driver

    def select_group(self):
        sleep(sleep_magic)
        driver = self.driver
        option_first = driver.find_elements(By.TAG_NAME, "mat-list-option")
        option = [i for i in option_first if i.text == "Research"][0]
        option.click()
        sleep(sleep_magic)
        groups = driver.find_elements(By.TAG_NAME, "mat-list-option")
        group = [i for i in groups if i.text == self.group][0]
        group.click()
        driver.find_element(By.ID, "firm-input-field").send_keys("CICC")
        sleep(sleep_magic)
        driver.find_element(By.TAG_NAME, "mat-option").click()
        sleep(sleep_magic)
        driver.execute_script(
            'document.getElementsByClassName("rating-selection large")[0].getElementsByTagName("mat-icon")[4].click();'
        )
        sleep(1)
        buttons = driver.find_elements(By.TAG_NAME, "button")
        button = [i for i in buttons if i.text == "Save"][0]
        button.click()
        logging.info(f"Group {self.group} selected")
        self.driver = driver

    def select_analyst(self, names: list[str]):
        sleep(sleep_magic)
        driver = self.driver
        companies = driver.find_elements(By.TAG_NAME, "mat-list-item")
        company = [
            i for i in companies if "China International Capital Corp" in i.text
        ][0]
        company.click()
        sleep(sleep_magic)
        for num, analyst_name in enumerate(names):
            driver.execute_script(
                f"""
            let analyst = [...document.getElementsByTagName("mat-list-item")].filter(e=>e.innerText.includes("{analyst_name}"))[0];
            analyst.getElementsByTagName("mat-icon")[4-{num}].click()
            """
            )
            logging.info(f"{analyst_name} has selected with {5-num} stars")
        driver.find_element(By.ID, "save-firm-vote-btn").click()
        self.driver = driver

    def screenshot(self):
        path = Path("screenshots")
        group_path = path / self.group
        if not group_path.exists():
            group_path.mkdir()
        sleep(sleep_magic)
        driver = self.driver
        driver.find_elements(By.TAG_NAME, "a")[2].click()
        sleep(sleep_magic)
        driver.find_element(By.CLASS_NAME, "node-title").click()
        sleep(sleep_magic * 2)
        groups = driver.find_elements(By.CLASS_NAME, "second-step-node")
        group = [i for i in groups if self.group in i.text][0]
        group.click()
        sleep(sleep_magic * 2)
        firms = driver.find_elements(By.CLASS_NAME, "firm-node")
        firm = [i for i in firms if "China International Capital Corp." in i.text][0]
        firm.click()
        driver.find_element(By.CLASS_NAME, "voting-property-node").click()
        driver.execute_script(
            f"""
        [...document.getElementsByClassName("second-step-node")].filter(e=>!e.innerText.includes("{self.group}")).map(e=>e.parentNode.removeChild(e))
        """
        )
        driver.get_screenshot_as_file(str(group_path / (self.user + ".png")))
        logging.info(f"take screenshot of {self.user} in {self.group}")
        driver.quit()

    def run(self, analysts: list[str]):
        self.login()
        self.select_group()
        self.select_analyst(list(analysts))
        self.screenshot()

if __name__ == "__main__":
    df = pd.read_excel("II投票.xlsx")
    data = df[["邮箱", "密码"]].dropna().to_dict("records")

    for i in data:
        user = i["邮箱"]
        password = i["密码"]
        option = selenium.webdriver.ChromeOptions()
        try:
            ip = PROXIES.pop()
            option.add_argument(f"--proxy-server={ip}")
            fv = FVote(user, password, GROUP, option)
            fv.run(ANALYSTS)
        except IndexError:
            logging.error("No proxy available!")
            fv = FVote(user, password, GROUP, option)
            fv.run(ANALYSTS)
            break
        sleep(sleep_magic)

# %%
