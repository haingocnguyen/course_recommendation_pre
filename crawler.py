import time
import xlwt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Set up Chrome WebDriver
chrome_options = Options()
path = r"D:\App_exe\chromedriver-win64_2\chromedriver-win64\chromedriver.exe"  # Update with your chromedriver path
ser = Service(path)
browser = webdriver.Chrome(service=ser)

# Open the Coursera URL for standalone courses (first page)
browser.get("https://www.coursera.org/courses?productTypeDescription=Courses&page=16&sortBy=BEST_MATCH")
time.sleep(5)

# Initialize Excel workbook and sheet
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Course Data')

# Define column titles for course data
columns = ["Title","Course Link", "Rating", "Difficulty", "Duration", "Provider", "Keywords", 
           "Financial Aid", "Enrollment Count", "Number of Modules", "Post-Course Evaluation", "Course Description"]
for i, column in enumerate(columns):
    sheet.write(0, i, column)

# Function to get element text or return default
def get_element_text_or_default(xpath, default="N/A"):
    try:
        return browser.find_element(By.XPATH, xpath).text
    except Exception as e:
        print(f"Error getting element: {e}")
        return default

# Function to check for financial aid availability
def check_financial_aid():
    try:
        # Look for the financial aid button
        finaid_button = browser.find_element(By.XPATH, '//p[@class="caption-text"]//button[contains(@class, "finaid-link")]')
        return "Available" if finaid_button.is_displayed() else "Not Available"
    except Exception:
        return "Not Available"

# Function to extract the number of enrolled students
def get_enrollment_count():
    try:
        enrollment_text = browser.find_element(By.XPATH, '//p[@class="css-4s48ix"]//strong/span').text
        return enrollment_text.replace(",", "")  # Remove commas for consistency
    except Exception:
        return "N/A"

# Function to extract number of modules
def get_number_of_modules():
    try:
        modules_text = browser.find_element(By.XPATH, '//a[@href="#modules"]').text
        return modules_text.split()[0]  # Get the number part (e.g., '5 modules')
    except Exception:
        return "N/A"

# Function to extract post-course evaluation (e.g., 98% liked it)
def get_post_course_evaluation():
    try:
        evaluation_text = browser.find_element(By.XPATH, '//span[@class="css-fk6qfz"]').text
        return evaluation_text
    except Exception:
        return "N/A"

# Function to extract course description
def get_course_description():
    try:
        outcomes = []
        items = browser.find_elements(By.XPATH, '//div[contains(@class, "rc-CML unified-CML")]//p/span/span')
        for item in items:
            outcomes.append(item.text)
        return "; ".join(outcomes)  # Join learning outcomes into a single string
    except Exception:
        return "N/A"

# Initialize row index for Excel
row_index = 1

# Collect all course URLs before navigating
course_elements = browser.find_elements(By.XPATH, '//a[@data-click-key="seo_entity_page.search.click.search_card"]')
course_urls = [course.get_attribute('href') for course in course_elements]

def scrape_course_data():
    global row_index
    # Loop through all the stored course URLs
    for course_link in course_urls:
        try:
            if course_link:
                # Navigate directly to the course page
                browser.get(course_link)
                time.sleep(5)  # Wait for the course page to load

                # Extract course details from the detailed page
                # Extract course link
                #full_course_link = course_link
                title = get_element_text_or_default('//h1[@data-e2e="hero-title"]', "N/A")
                rating = get_element_text_or_default('//div[contains(@class, "css-h1jogs") and contains(@class, "cds-119")]', "N/A")
                difficulty = get_element_text_or_default('.//div[@class="css-dwgey1"]//div[contains(@class, "css-139h6xi")]//div[contains(@class, "css-fk6qfz")]', "N/A")
                duration = get_element_text_or_default('.//div[@class="css-fw9ih3"]/div[1]', "N/A")
                provider = get_element_text_or_default('//div[contains(@class, "partner-name")]//span', "N/A")
                keywords = get_element_text_or_default('//ul[contains(@class, "css-yk0mzy")]/li/span[contains(@class, "css-o5tswl")]', "N/A")
                # Additional data points
                financial_aid = check_financial_aid()
                enrollment_count = get_enrollment_count()
                number_of_modules = get_number_of_modules()
                post_course_evaluation = get_post_course_evaluation()
                course_description = get_course_description()

                # Write data to Excel
                data = [title, course_link, rating, difficulty, duration, provider, keywords, 
                        financial_aid, enrollment_count, number_of_modules, post_course_evaluation, course_description]
                for i, value in enumerate(data):
                    sheet.write(row_index, i, value)

                row_index += 1  # Move to the next row

        except Exception as e:
            print(f"Error extracting data from course: {e}")

# Scrape courses from the first page only
scrape_course_data()

# Save the Excel workbook
workbook.save(r'D:\Thesis\data\des_main\page16.xls')  # Update the path to save your Excel file

# Close the browser after scraping is done
browser.quit()
