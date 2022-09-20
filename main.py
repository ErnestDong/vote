import logging
from pathlib import Path
from time import sleep

import pandas as pd
import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
sleep_magic = 5
df = (
    pd.read_excel(
        "II投票.xlsx",
        sheet_name="其他组",
        index_col="行业",
        usecols=["行业", "投票1", "投票2", "投票3", "投票4", "投票5"],
    )
    .dropna(how="all")
    .to_dict(orient="index")
)
translate = {
    "技术硬件": "Technology Hardware",
    "电信": "Telecommunications",
    "农业": "Agriculture",
    "医药": "Healthcare, Pharmaceuticals & Biotechnology",
    "化工": "Chemicals",
    "石油": "Oil & Gas",
    "公用事业与新能源": "Public Utilities & Alternative Energy",
    "非必须消费品": "Consumer Discretionary",
    "日用消费品": "Consumer Staples",
    "策略": "Strategy",
}
vote_data = {
    translate[i]: [i for i in df[i].values() if isinstance(i, str)] for i in df
}


class FVote:
    def __init__(self, user: str, password: str, option=None) -> None:
        self.user = user
        self.password = password
        if option is None:
            option = selenium.webdriver.ChromeOptions()
        self.driver = selenium.webdriver.Chrome(options=option)
        # self.group = groupname

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

    def select_group(self, groupname):
        sleep(sleep_magic)
        driver = self.driver
        option_first = driver.find_elements(By.TAG_NAME, "mat-list-option")
        option = [i for i in option_first if i.text == "Research"][0]
        option.click()
        sleep(sleep_magic)
        groups = driver.find_elements(By.TAG_NAME, "mat-list-option")
        group = [i for i in groups if i.text == groupname][0]
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
        logging.info(f"Group {groupname} selected")
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
        sleep(sleep_magic)
        # driver.find_elements(By.TAG_NAME, "r2-breadcrumb")[1].click()
        self.driver = driver

    def screenshot(self, groupname):
        path = Path("screenshots")
        group_path = path / groupname
        if not group_path.exists():
            group_path.mkdir()
        sleep(sleep_magic)
        driver = self.driver
        driver.find_elements(By.TAG_NAME, "a")[2].click()
        sleep(sleep_magic)
        driver.find_elements(By.CLASS_NAME, "node-title")[-1].click()
        sleep(sleep_magic * 2)
        groups = driver.find_elements(By.CLASS_NAME, "second-step-node")
        group = [i for i in groups if groupname in i.text][0]
        group.click()
        sleep(sleep_magic * 2)
        firms = driver.find_elements(By.CLASS_NAME, "firm-node")
        firm = [i for i in firms if "China International Capital Corp." in i.text][0]
        firm.click()
        driver.find_element(By.CLASS_NAME, "voting-property-node").click()
        driver.execute_script(
            f"""
[...document.getElementsByClassName("second-step-node")].filter(e=>!e.innerText.includes("{groupname}")).map(e=>e.parentNode.removeChild(e))
        """
        )
        driver.get_screenshot_as_file(str(group_path / (self.user + ".png")))
        logging.info(f"take screenshot of {self.user} in {groupname}")

    def run(self, data: dict):
        self.login()
        for groupname, analysts in data.items():
            if (Path("screenshots") / groupname / (self.user + ".png")).exists():
                logging.info(f"{self.user} in {groupname} has already voted")
                continue
            self.select_group(groupname)
            self.select_analyst(analysts)
            self.screenshot(groupname)
            self.driver.find_elements(By.TAG_NAME, "a")[1].click()
            sleep(sleep_magic)
        self.driver.quit()


if __name__ == "__main__":
    df = pd.read_excel("II投票.xlsx", sheet_name="我的")
    data = df[["邮箱", "密码"]].dropna().to_dict("records")

    for i in data:
        user = i["邮箱"]
        password = i["密码"]
        option = selenium.webdriver.ChromeOptions()
        fv = FVote(user, password, option)
        fv.run(vote_data)
        sleep(sleep_magic)
