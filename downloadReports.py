from playwright.async_api import Playwright, async_playwright, Error
import asyncio
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
from encryptPass import password, user, url, company
import os


async def safe_goto(page, url, max_retries=10, delay=0.5):
    """Function to safely go to a specific url using playwright
    This fixes the issue where you may not connect to the address
    on the first try. This will try about 10 times until you can connect"""
    for attempt in range(1, max_retries + 1):
        try:
            response = await page.goto(url)
            print("Navigation successful!")
            return response  # or perform further actions with the response
        except Error as e:
            # Check if the error message indicates a DNS resolution issue
            if "net::ERR_NAME_NOT_RESOLVED" in str(e):
                print(
                    f"Attempt {attempt} failed with error: {e}. Retrying in {delay} second(s)..."
                )
                await asyncio.sleep(delay)
            else:
                # If it's a different error, re-raise it
                raise
    # If all attempts fail, you can choose to raise an exception or handle it gracefully
    raise Exception(f"Failed to navigate to {url} after {max_retries} attempts.")


async def uncheck(page, locationName):
    """Function to uncheck a checkbox element on E2
    This function first tries to find the element using the locator function
    if that does not work it will use the get_by_text function"""
    checkbox = page.locator(locationName)
    if await checkbox.count() == 0:
        checkbox = page.get_by_text(locationName)
    if await checkbox.is_checked():
        await checkbox.click()


async def login_e2(page):
    """Function to log into E2 and this will use the seperate module encryptPass
    for checking if a user name and password exists if not then it will
    prompt the user to create one for their E2 account"""
    await page.get_by_role("textbox", name="").fill(user())
    await page.get_by_role("textbox", name="").fill(password())
    await page.get_by_role("button", name=" LOGIN").click()
    await page.locator("#company").select_option(company())
    await page.get_by_role("link", name="Main").click()
    # handle if the "already logged in" error happens
    # logout other session in order to login to e2
    try:
        await page.get_by_text("This user is already logged").click(timeout=3000)
        await page.get_by_role("button", name=" Yes").click()
    except Error as e:
        if "Timeout" in str(e):
            pass


async def logout_e2(page):
    """Function to log out of E2, right now it only works for username max and
    this needs to be made more general for everyone"""
    await page.get_by_role("link", name=" Welcome, MAX KRANKER ").click()
    await page.locator("#navbar-container").get_by_text("Logout").click()


async def download_csv(page, data_folder):
    try:
        async with page.expect_download(timeout=120000) as download_info:
            await page.get_by_role("button", name=" Generate Report ").click()
            await page.get_by_role("link", name="Export as... ").click()
            await page.get_by_role("link", name="CSV").click()
        download = await download_info.value
        downloadName = (
            os.path.abspath(".") + f"/{data_folder}/" + download.suggested_filename
        )
        await download.save_as(downloadName)
        await page.get_by_role("button", name=" Close").click()
    except Error as e:
        print(e)


async def setReportDateRange(page, nthchild, begDate, endDate):
    begDateStr = begDate.strftime("%m/%d/%Y")
    endDateStr = endDate.strftime("%m/%d/%Y")
    await page.locator('input[name="rptDate"]').click()
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_start_input > .input-mini"
    ).click()
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_start_input > .input-mini"
    ).fill("")
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_start_input > .input-mini"
    ).fill(begDateStr)
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_end_input > .input-mini"
    ).click()
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_end_input > .input-mini"
    ).fill("")
    await page.locator(
        f"div:nth-child({nthchild}) > .ranges > .range_inputs > .daterangepicker_end_input > .input-mini"
    ).fill(endDateStr)
    await page.get_by_role("button", name="Apply").click()


async def download_job_schedule(page):
    await page.get_by_role("button", name="").click()
    await page.get_by_role("link", name=" Orders ").click()
    await page.get_by_role("link", name="Job Schedule").click()
    checkboxList = {
        "#cg_rptPartNumberExclude span",
        "#cg_rptPartNumberInclude span",
        "#cg_rptPartDescriptionExclude span",
        "#cg_rptPartDescriptionInclude span",
        "#cg_rptJobNotesExclude span",
        "#cg_rptJobNotesInclude span",
        "#cg_rptCustPONumberExclude span",
        "#cg_rptCustPONumberInclude span",
        "#cg_rptOrderNumberExclude span",
        "#cg_rptOrderNumberInclude span",
        "#cg_rptProductCodeExclude span",
        "#cg_rptProductCodeInclude span",
        "#cg_rptCustomerCodeExclude span",
        "#cg_rptCustomerCodeInclude span",
        "#cg_rptWorkCodeExclude span",
        "#cg_rptWorkCodeInclude span",
        "#cg_rptSalesmanExclude span",
        "#cg_rptSalesmanInclude span",
        "#cg_rptCurrentWorkCenterExclude span",
        "#cg_rptCurrentWorkCenterInclude span",
        "#cg_rptPriorityExclude span",
        "#cg_rptPriorityInclude span",
        "#cg_rptDateInclude span",
        "#cg_rptReadyToShipStatusInclude span",
        "#cg_rptInProcessStatusInclude span",
        "#cg_rptReleaseTypesInclude span",
        "#cg_rptJobsOnHoldInclude span",
        "Show Prior Shipment",
    }
    for name in checkboxList:
        await uncheck(page, name)

    await page.get_by_text("Due Date").click()
    begDate = datetime.strptime("12/01/2023", "%m/%d/%Y")
    endDate = datetime.strptime("1/1/2027", "%m/%d/%Y")
    await setReportDateRange(page, 72, begDate, endDate)
    await download_csv(page, "csv_files")
    """
    dateDivider = 1
    while dateDivider != 20:
        tempBegDate = begDate
        dateDifference = abs((endDate - begDate).days)
        dateDelta = int(dateDifference / dateDivider)
        dateDeltaRemainder = dateDelta + (dateDifference % dateDivider)
        for i in range(1, dateDivider, 1):
            tempEndDate = tempBegDate + timedelta(days=dateDelta - 1)
            try:
                await setReportDateRange(page, 72, tempBegDate, tempEndDate)
                await download_csv(page, i)
            except Error as e:
                print(e)
                continue
            tempBegDate = tempEndDate + timedelta(days=1)
        print("Download Success")
        break"""


async def download_order_entry_detail(page):
    await page.get_by_role("button", name="").click()
    await page.get_by_role("link", name=" Orders ").click()
    await page.get_by_role("link", name="Order Entry Summary").click()
    checkboxList = {
        "#cg_rptPartNumberExclude span",
        "#cg_rptPartNumberInclude span",
        "#cg_rptPartDescriptionExclude span",
        "#cg_rptPartDescriptionInclude span",
        "#cg_rptJobNotesExclude span",
        "#cg_rptJobNotesInclude span",
        "#cg_rptCustPoNumberExclude span",
        "#cg_rptCustPoNumberInclude span",
        "#cg_rptOrderNumberExclude span",
        "#cg_rptOrderNumberInclude span",
        "#cg_rptProductCodeExclude span",
        "#cg_rptProductCodeInclude span",
        "#cg_rptCustomerCodeExclude span",
        "#cg_rptCustomerCodeInclude span",
        "#cg_rptWorkCodeExclude span",
        "#cg_rptWorkCodeInclude span",
        "#cg_rptSalesmanExclude span",
        "#cg_rptSalesmanInclude span",
        "#cg_rptPriorityExclude span",
        "#cg_rptPriorityInclude span",
        "#cg_rptDateInclude span",
        "Detail",
        "Customer Breakdown",
        "Part Number Breakdown",
        "Work Code Breakdown",
        "Sales ID Breakdown",
        "Product Code Breakdown",
        "Breakdowns as Graphs",
    }

    for name in checkboxList:
        await uncheck(page, name)
    await page.get_by_text("Detail").click()
    await page.get_by_text("Date Due").click()
    begDate = datetime.strptime("12/01/2023", "%m/%d/%Y")
    endDate = datetime.strptime("1/1/2027", "%m/%d/%Y")
    await setReportDateRange(page, 71, begDate, endDate)
    await download_csv(page, "csv_files")


async def download_purchase_order_detail(page):
    await page.get_by_role("button", name="").click()
    await page.get_by_role("link", name=" Purchasing ").click()
    await page.get_by_role("link", name="Purchase Order Summary").click()

    checkboxList = {
        "#cg_rptPartNumberExclude span",
        "#cg_rptPartNumberInclude span",
        "#cg_rptPartDescriptionExclude span",
        "#cg_rptPartDescriptionInclude span",
        "#cg_rptCommentsExclude span",
        "#cg_rptCommentsInclude span",
        "#cg_rptJobNumberExclude span",
        "#cg_rptJobNumberInclude span",
        "#cg_rptVendorCodeExclude span",
        "#cg_rptVendorCodeInclude span",
        "#cg_rptVendorTypeExclude span",
        "#cg_rptVendorTypeInclude span",
        "#cg_rptPurchasedByExclude span",
        "#cg_rptPurchasedByInclude span",
        "#cg_rptGLCodeExclude span",
        "#cg_rptGLCodeInclude span",
        "#cg_rptCurrencyCodeExclude span",
        "#cg_rptCurrencyCodeInclude span",
        "#cg_rptStatusInclude span",
        "#cg_rptPOTypeInclude span",
        "#cg_rptLocationInclude span",
        "#cg_rptDateInclude span",
        "Detail",
        "Vendor Breakdown",
        "GL Code Breakdown",
    }

    for name in checkboxList:
        await uncheck(page, name)

    await page.get_by_text("Detail").click()
    await page.get_by_text("Date Entered").click()
    begDate = date.today() - relativedelta(years=2)
    endDate = date.today()
    await setReportDateRange(page, 72, begDate, endDate)
    await download_csv(page, "csv_files")


async def run(playwright: Playwright) -> None:
    browser = await playwright.chromium.launch(
        headless=False,
    )
    page = await browser.new_page()
    try:
        await safe_goto(page, url(), 10)
    except Exception as e:
        print(e)
    await login_e2(page)
    try:
        await download_purchase_order_detail(page)
        await download_order_entry_detail(page)
        await download_job_schedule(page)
        await logout_e2(page)
    except Error as e:
        print(e)
        await logout_e2(page)
    # ---------------------
    await browser.close()


async def main_playwright():
    async with async_playwright() as playwright:
        await run(playwright)


def download_reports():
    asyncio.run(main_playwright())
