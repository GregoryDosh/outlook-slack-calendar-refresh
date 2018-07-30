import datetime
import logging
import os
import time

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

_LOGGER = logging.getLogger(__name__)

class SeleniumOutlook(object):
    def __init__(self, email_address, userid, password):
        super(SeleniumOutlook, self).__init__()
        self.email_address = email_address
        self.userid = userid
        self.password = password
        self.driver = None
        self.logged_in = False

    def __enter__(self):
        _LOGGER.debug("Opening chromedriver")

        options = webdriver.ChromeOptions()
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--ignore-urlfetcher-cert-requests")

        capabilities = options.to_capabilities()

        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(5)
        self.driver.get("https://outlook.office.com/")
        try:
            # Input username
            self.driver.find_element_by_name("loginfmt").send_keys(self.email_address)
            # Click submit
            self.driver.find_element_by_xpath("//input[@type='submit']").click()
            time.sleep(8)
            # If finding the corp login, then do that
            try:
                # Input username
                self.driver.find_element_by_id("loginID").send_keys(self.userid)

                # Input password
                self.driver.find_element_by_id("pass").send_keys(self.password)

                # Click submit
                self.driver.find_element_by_xpath("//button[@type='submit']").click()
            except selenium.common.exceptions.NoSuchElementException:
                _LOGGER.debug("No login found, assuming SAML took over")
            self.driver.find_element_by_id("idSIButton9").click()
        except selenium.common.exceptions.NoSuchElementException:
            _LOGGER.debug("Should be logged in already?")
        # This is to wait for the page to load so we know we logged in...
        time.sleep(2)
        try:
            self.driver.find_element_by_class_name("ms-Icon--calendar")
            self.logged_in = True
        except selenium.common.exceptions.NoSuchElementException:
            _LOGGER.error("Login Failed")
            self.driver.close()
            raise
        _LOGGER.debug("Successful Login")
        return self

    def __exit__(self, *args, **kwargs):
        if self.driver:
            self.driver.close()

    def _hide_left_nav_bar(self):
        try:
            self.driver.find_element_by_xpath("//div[@style='position: absolute; top: 0px; right: auto; bottom: 0px; left: 0px; height: auto; width: 210px;']").click()
            self.driver.find_element_by_xpath("//button[@aria-label='collapse the navigation pane']").click()
        except selenium.common.exceptions.NoSuchElementException as e:
            _LOGGER.debug("Assuming navbar already hidden")

    def _send_offset_shortcuts(self, time_offset):
        while time_offset > 0:
            self.driver.find_element_by_tag_name('body').send_keys(Keys.SHIFT + Keys.RIGHT)
            time_offset -= 1
            time.sleep(0.5)
        while time_offset < 0:
            self.driver.find_element_by_tag_name('body').send_keys(Keys.SHIFT + Keys.LEFT)
            time_offset += 1
            time.sleep(0.5)

    def screenshot_day_calendar(self, picture_location, time_offset=0, width=960, height=1080):
        if self.logged_in:
            self.driver.set_window_size(width, height)
            self.driver.get("https://outlook.office.com/owa/?path=/calendar/view/Day")
            self._hide_left_nav_bar()
            self._send_offset_shortcuts(time_offset)
            time.sleep(2.5)
            self.driver.get_screenshot_as_file(os.path.expanduser(picture_location))
        else:
            _LOGGER.error("Can't screen shot when not logged in")

    def screenshot_week_calendar(self, picture_location, time_offset=0, width=1280, height=1080):
        if self.logged_in:
            self.driver.set_window_size(width, height)
            self.driver.get("https://outlook.office.com/owa/?path=/calendar/view/WorkWeek")
            self._hide_left_nav_bar()
            self._send_offset_shortcuts(time_offset)
            time.sleep(2.5)
            self.driver.get_screenshot_as_file(os.path.expanduser(picture_location))
        else:
            _LOGGER.error("Can't screen shot when not logged in")

class SlackCalendarUpload(object):
    def __init__(self, userid, password, slack_login_url, slack_private_message_url):
        super(SlackCalendarUpload, self).__init__()
        self.userid = userid
        self.password = password
        self.slack_login_url = slack_login_url
        self.slack_private_message_url = slack_private_message_url

        self.driver = None
        self.logged_in = False

    def __enter__(self):
        _LOGGER.debug("Opening chromedriver")
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(5)
        self.driver.get(self.slack_login_url)
        try:
            self.driver.find_element_by_xpath("//html//div[@class='card sign_in_sso_card larger_top_padding larger_bottom_padding']/a[1]").click()

            # Input username
            self.driver.find_element_by_id("loginID").send_keys(self.userid)

            # Input password
            self.driver.find_element_by_id("pass").send_keys(self.password)

            # Click submit
            self.driver.find_element_by_xpath("//button[@type='submit']").click()

        except selenium.common.exceptions.NoSuchElementException:
            _LOGGER.debug("Should be logged in already?")

        # Navigate to private messages
        self.driver.get(self.slack_private_message_url)

        # This is to wait for the page to load so we know we logged in...
        try:
            self.driver.find_element_by_xpath("//div[@class='ql-placeholder'][contains(text(),'Jot something down')]")
            self.logged_in = True
            time.sleep(5)
        except selenium.common.exceptions.NoSuchElementException:
            _LOGGER.error("Login Failed")
            self.driver.close()
            raise
        _LOGGER.debug("Successful Login")
        return self

    def __exit__(self, *args, **kwargs):
        if self.driver:
            self.driver.close()

    def remove_pictures(self):
        _LOGGER.debug("Removing all pictures from private messages")
        while True:
            try:
                _LOGGER.debug("Looking for images")
                elements = self.driver.find_elements_by_class_name("c-message__file--image")
                if len(elements) == 0:
                    _LOGGER.debug("No pictures found")
                    break
                _LOGGER.debug("Found {} images".format(len(elements)))
                for el in reversed(elements):
                    _LOGGER.debug("removing an image")
                    el.click()
                    self.driver.find_element_by_tag_name('body').send_keys(Keys.TAB)
                    el.find_elements_by_class_name('c-file__action_button')[2].click()
                    for _ in range(6):
                        self.driver.find_element_by_tag_name('body').send_keys(Keys.TAB)

                    self.driver.find_element_by_class_name('c-menu_item__li--highlighted').click()
                    time.sleep(1)

                    self.driver.find_element_by_xpath("//button[@type='button'][contains(text(),'Yes, delete this file')]").click()
                    time.sleep(1)
                _LOGGER.debug("Double checking images")
            except selenium.common.exceptions.NoSuchElementException:
                _LOGGER.debug("All pictures deleted?")
                break


    def upload_file(self, filename, title, maxWait=15):
        _LOGGER.debug("Uploading {}".format(filename))
        self.driver.find_element_by_id("primary_file_button").click()
        self.driver.find_element_by_xpath("//li[@data-which='choose']").click()
        file_upload = self.driver.find_element_by_xpath("//input[@type='file']")
        file_upload.send_keys(os.path.expanduser(filename))
        self.driver.find_element_by_class_name("p-file-upload_dialog__preview_file_name_edit").send_keys(Keys.ENTER)
        title_el = self.driver.find_element_by_id("file-upload-name")
        title_el.send_keys(title)
        time.sleep(1.25)
        self.driver.find_element_by_class_name("c-dialog__go").click()
        time.sleep(0.25)
        while maxWait > 0:
            if 'Processing uploaded file' in self.driver.page_source:
                _LOGGER.debug("waiting on upload {}...".format(maxWait))
                time.sleep(1)
                maxWait -= 1
            else:
                _LOGGER.debug("File uploaded succesfully")
                break
        if maxWait == 0:
            _LOGGER.warn("Pictures may not have uploaded completely")

if __name__ == '__main__':
    with SeleniumOutlook(email_address=os.environ['OUTLOOK_EMAIL'],
                         userid=os.environ['USER_ID'],
                         password=os.environ['USER_PASS'],
                         ) as ol:

        today = datetime.datetime.now()
        if today.strftime("%A") != "Friday":
            ol.screenshot_day_calendar("/tmp/today.png")
            ol.screenshot_day_calendar("/tmp/tomorrow.png", time_offset=1)
            ol.screenshot_week_calendar("/tmp/this_week.png")
            ol.screenshot_week_calendar("/tmp/next_week.png", time_offset=1)
        else:
            ol.screenshot_day_calendar("/tmp/today.png", time_offset=3)
            ol.screenshot_day_calendar("/tmp/tomorrow.png", time_offset=4)
            ol.screenshot_week_calendar("/tmp/this_week.png", time_offset=1)
            ol.screenshot_week_calendar("/tmp/next_week.png", time_offset=2)

    with SlackCalendarUpload(slack_private_message_url=os.environ['SLACK_PRIVATE_MESSAGE_URL'],
                             slack_login_url=os.environ['SLACK_LOGIN_URL'],
                             userid=os.environ['USER_ID'],
                             password=os.environ['USER_PASS'],
                             ) as sl:
        sl.remove_pictures()
        sl.upload_file("/tmp/today.png", "Today")
        sl.upload_file("/tmp/tomorrow.png", "Tomorrow")
        sl.upload_file("/tmp/this_week.png", "This Week")
        sl.upload_file("/tmp/next_week.png", "Next Week")
    _LOGGER.info("Upload complete {}".format(time.strftime("%a, %d %b %Y %H:%M:%S", time.localtime())))
