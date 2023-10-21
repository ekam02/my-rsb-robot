from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


@task
def robot_spare_bin_python():
    """Introduce los datos de ventas de la semana y los exporta en formato PDF"""
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()


def open_the_intranet_website():
    """Navega a la URL indicada"""
    browser.goto("https://robotsparebinindustries.com/")


def log_in():
    """Rellena el formulario de inicio de sesión y pulsa el botón 'Log in'"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")


def fill_and_submit_sales_form(sales_rep):
    """Rellena los datos de venta y pulsa el botón 'Submit'."""
    page = browser.page()
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")


def download_excel_file():
    """Descarga el archivo Excel de la URL indicada"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)


def fill_form_with_excel_data():
    """Leer datos de excel y rellenar el formulario de ventas"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()
    for row in worksheet:
        fill_and_submit_sales_form(row)


def collect_results():
    """Haz una captura de pantalla de la página"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")


def log_out():
    """Pulsa el botón 'Log out'"""
    page = browser.page()  
    page.click("text=Log out")


def export_as_pdf():
    """Exportar los datos a un archivo pdf"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
