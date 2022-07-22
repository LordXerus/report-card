import pathlib
import re
import time
from itertools import islice

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait, Select
from webdriver_manager.firefox import GeckoDriverManager


def get_all_years():
    year_element: Select = Select(
        WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable(
            (By.XPATH, "//select[@id='schYear']")
        ))
    )

    return [option.text for option in year_element.options]


def switch_to_year(year: str):
    year_element: Select = Select(
        WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable(
            (By.XPATH, "//select[@id='schYear']")
        ))
    )

    text = next(filter(
        lambda y: y.endswith(year),
        (o.text for o in year_element.options)
    ))

    year_element.select_by_visible_text(text)

    time.sleep(1)

    return text


def print_years(years, row_length=5):
    year_iter = iter(years)
    while True:
        group = list(islice(year_iter, row_length))
        if not group:
            break
        print(*group)


def prompt_range():
    print('I have data for:')
    available_years = list(year[-2:] for year in get_all_years())
    print_years(available_years)

    print()
    print("Please input years:")
    print('-> Comma-separated list')
    print('-> xx to include year')
    print('-> xx-xx to include range')
    print('-> ~xx to exclude range')
    print('-> Blank answer/trailing comma will download latest')

    answer = input('> ')
    available_years = set(available_years)
    included = set()

    for year_ in answer.split(','):
        year = year_.strip()

        range_match = re.fullmatch(r'(\d{1,2})-(\d{1,2})', year)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))

            for ans in range(start, end + 1):
                result = str(ans)
                if result in available_years:
                    included.add(result)

            continue

        not_match = re.fullmatch(r'~(\d{1,2})', year)
        if not_match:
            number = not_match.group(1)
            included.discard(number)
            continue

        single_match = re.fullmatch(r'\d{1,2}', year)
        if single_match:
            number = single_match.group(0)
            if number in available_years:
                included.add(number)
            else:
                print(f'WARNING: {number} not found in available data.')

        if not year:
            latest = max(available_years, key=int)
            included.add(latest)
            print(f'Downloading latest({latest})')

    included = list(included)
    included.sort()

    print()
    print('Included years:')
    print_years(included)

    return included


if __name__ == '__main__':
    home_path = pathlib.Path.cwd() / 'Report Card Downloader'
    home_path.resolve()
    if not home_path.exists():
        home_path.mkdir()

    tmp_path = home_path / 'tmp'
    if not tmp_path.exists():
        tmp_path.mkdir()

    firefoxOptions = Options()

    firefoxOptions.set_preference("browser.download.folderList", 2)
    firefoxOptions.set_preference("browser.download.manager.showWhenStarting", False)
    firefoxOptions.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")

    firefoxOptions.set_preference("pdfjs.disabled", True)
    PDFJS_DISABLED = True

    print(f'Home Path is: {home_path}')
    print(f'Download Path is: {tmp_path}')
    firefoxOptions.set_preference("browser.download.dir", str(tmp_path))
    service = FirefoxService(executable_path=GeckoDriverManager().install())

    driver = webdriver.Firefox(options=firefoxOptions, service=service)

    driver.get('https://schoolzone.epsb.ca/cf/index.cfm')

    USERNAME = input('Please input Username:')
    PASSWORD = input('Please input Password:')

    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    id_element = WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable(
        (By.XPATH, "//input[@id='userID']")
    ))

    id_element.send_keys(USERNAME)

    pass_element = WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable(
        (By.XPATH, "//input[@id='loginPassword']")
    ))

    pass_element.send_keys(PASSWORD)
    pass_element.submit()

    time.sleep(2)

    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # ++LOGGED IN++

    driver.get('https://schoolzone.epsb.ca/cf/profile/progressInterim/index.cfm')

    regex_bool = re.compile(r'true|false')

    for year in prompt_range():
        full_year = switch_to_year(year)
        print(f'Downloading {full_year}')

        time.sleep(1)

        table: WebElement = WebDriverWait(driver, 10).until(expected_conditions.element_to_be_clickable(
            (By.XPATH, "//table[@id='reportsTable']")
        ))

        tbody = table.find_element(By.XPATH, './tbody')

        for tr in tbody.find_elements(By.XPATH, './tr'):

            a = tr.find_element(By.XPATH, './td[4]/a')
            href = a.get_attribute('href')

            params_match = re.fullmatch(
                r"javascript:DoOpenWindow\('([^']*)'(?:,'(?:true|false)'){2}\)",
                href
            )

            if not params_match:
                raise RuntimeError('Cannot find download parameters')

            params = params_match.group(1)

            if not globals().get('PDFJS_DISABLED', None):
                window_handles = set(driver.window_handles)
                if len(window_handles) > 1:
                    raise RuntimeError("ERROR: More than 1 window detected. Window didn't close.")

            print(f'opening {params}')
            # Is ViewReport the window name?
            driver.execute_script(
                f'''
                window.open(
                    "Launch.cfm?{params}","ViewReport",
                    "toolbar=no,status=no,location=no,directories=no,copyhistory=no,height=450,width=600,resizable=yes"
                )
                '''
            )

            WebDriverWait(driver, 5).until(expected_conditions.number_of_windows_to_be(2))
            time.sleep(1)
            if not globals().get('PDFJS_DISABLED', None):
                opened_window_handles = set(driver.window_handles)
                new_window = next(iter(
                    opened_window_handles - window_handles
                ))

                old_handle = driver.current_window_handle

                driver.switch_to.window(new_window)

                download_button: WebElement = WebDriverWait(driver, 5).until(
                    expected_conditions.element_to_be_clickable(
                        (By.XPATH, "//button[@id='download']")
                    ))

                download_button.click()

                time.sleep(1)

                driver.close()
                driver.switch_to.window(old_handle)

        year_path = home_path / full_year
        if not year_path.exists():
            year_path.mkdir()

        for file in tmp_path.iterdir():
            file.rename(year_path / file.name)

    input('Press enter to kill browser...')
    driver.quit()
