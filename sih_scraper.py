from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time

URL = "https://www.sih.gov.in/sih2026PS"

OUTPUT_FILE = "SIH_2026_226_Problem_Statements.xlsx"

EXPECTED_TOTAL = 226

options = webdriver.ChromeOptions()

options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 30)

try:

    print()
    print("==============================================")
    print(" SMART INDIA HACKATHON 2026 SCRAPER")
    print("==============================================")
    print()

    print("Opening SIH website...")

    driver.get(URL)

    wait.until(
        EC.presence_of_element_located(
            (By.ID, "dataTablePS")
        )
    )

    print("Main table loaded.")

    try:

        dropdown = wait.until(
            EC.presence_of_element_located(
                (By.NAME, "dataTablePS_length")
            )
        )

        Select(dropdown).select_by_value("100")

        time.sleep(2)

        print("Displaying 100 problems per page.")

    except Exception as e:

        print(
            "Could not set 100 rows:",
            e
        )

    MAIN_ROWS = (
        By.XPATH,
        "//table[@id='dataTablePS']/tbody/tr"
    )

    all_data = []

    page_number = 1

    while True:

        print()
        print("----------------------------------------------")
        print(f"PROCESSING PAGE {page_number}")
        print("----------------------------------------------")

        wait.until(
            EC.presence_of_element_located(
                MAIN_ROWS
            )
        )

        rows = driver.find_elements(
            *MAIN_ROWS
        )

        print(
            "Main problem rows:",
            len(rows)
        )

        for row_number in range(len(rows)):

            try:

                rows = driver.find_elements(
                    *MAIN_ROWS
                )

                row = rows[row_number]

                cells = row.find_elements(
                    By.XPATH,
                    "./td"
                )

                if len(cells) != 8:

                    print(
                        f"WARNING: Problem row "
                        f"{row_number + 1} has "
                        f"{len(cells)} cells"
                    )

                    continue

                values = []

                for cell in cells:

                    text = cell.get_attribute(
                        "textContent"
                    )

                    if text is None:

                        text = ""

                    text = text.replace(
                        "\xa0",
                        " "
                    )

                    text = " ".join(
                        text.split()
                    )

                    values.append(
                        text.strip()
                    )

                all_data.append(values)

                print(
                    f"{values[0]:>3} | "
                    f"{values[4]:<10} | "
                    f"{values[2][:75]}"
                )

            except Exception as e:

                print(
                    f"ERROR processing row "
                    f"{row_number + 1}: {e}"
                )

        next_button = driver.find_element(
            By.ID,
            "dataTablePS_next"
        )

        next_class = next_button.get_attribute(
            "class"
        )

        if "disabled" in next_class:

            print()
            print("Reached final page.")

            break

        current_rows = driver.find_elements(
            *MAIN_ROWS
        )

        if not current_rows:

            break

        old_first_row = current_rows[0]

        print("Moving to next page...")

        driver.execute_script(
            "arguments[0].click();",
            next_button
        )

        try:

            wait.until(
                EC.staleness_of(
                    old_first_row
                )
            )

        except:

            time.sleep(1)

        page_number += 1

    print()
    print("Removing duplicates...")

    unique_data = []

    seen = set()

    for row in all_data:

        ps_number = row[4]

        if ps_number not in seen:

            seen.add(
                ps_number
            )

            unique_data.append(
                row
            )

    all_data = unique_data

    print()
    print("Creating Excel file...")

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "SIH 2026 Problems"

    headers = [
        "S.No.",
        "Organization",
        "Problem Statement Title",
        "Category",
        "PS Number",
        "Submitted Idea(s) Count",
        "Theme",
        "Deadline for Idea Submission"
    ]

    sheet.append(
        headers
    )

    for row in all_data:

        sheet.append(
            row
        )

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = sheet.dimensions

    sheet.column_dimensions["A"].width = 10

    sheet.column_dimensions["B"].width = 40

    sheet.column_dimensions["C"].width = 70

    sheet.column_dimensions["D"].width = 15

    sheet.column_dimensions["E"].width = 15

    sheet.column_dimensions["F"].width = 25

    sheet.column_dimensions["G"].width = 30

    sheet.column_dimensions["H"].width = 30

    workbook.save(
        OUTPUT_FILE
    )

    print()
    print("==============================================")
    print(" SCRAPING FINISHED")
    print("==============================================")

    print(
        "Problems collected:",
        len(all_data)
    )

    print(
        "Expected:",
        EXPECTED_TOTAL
    )

    print(
        "Excel file:",
        OUTPUT_FILE
    )

    if len(all_data) == EXPECTED_TOTAL:

        print()
        print(
            "SUCCESS! All 226 problem statements collected."
        )

    else:

        print()
        print(
            "WARNING: Expected 226 but got",
            len(all_data)
        )

finally:

    driver.quit()
