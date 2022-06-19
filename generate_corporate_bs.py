from selenium import webdriver
from outlook_utilities import launch_outlook_api, send_email, format_email_recipients_from_list
from datetime import datetime

EMAIL_RECIPIENTS = []

BS_GENERATOR_URL = 'https://www.atrixnet.com/bs-generator.html'

def get_bs(bs_generator_url=BS_GENERATOR_URL):

    # Create webdriver and go to corporate BS website
    driver = webdriver.Chrome('Y:\\chromedriver.exe')
    driver.get(bs_generator_url)

    # Find and click button that generates BS
    generate_button = driver.find_element_by_xpath('//form[1]//input[@type="button"]')
    generate_button.click()

    # Grab generated text
    text_box = driver.find_element_by_id('bullshit')
    bs = text_box.get_attribute('value')

    driver.close()

    return bs

def get_email_subject():

    today = datetime.today()

    return f'Objective for {today.strftime("%Y/%m/%d")}'



if __name__ == 'main':

    outlook, outlook_api = launch_outlook_api()

    bs = get_bs()

    # Send an email with today's drivel
    send_email(
        outlook_session=outlook,
        subject=get_email_subject(),
        to=format_email_recipients_from_list(EMAIL_RECIPIENTS),
        body_html=f'<h4>{bs}</h4><br><br><p>Powered by {BS_GENERATOR_URL}</p>'
    )
