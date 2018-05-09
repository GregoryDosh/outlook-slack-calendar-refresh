# Why?
I don't know.  I was looking to play around with Selenium and so I made up a contrived example of something I wanted to accomplish.  I would like to look at my work's outlook calendar without having to sync the calendar on my phone or install any work related software.  Using slack I can login to our secure environment and still see my attachments without having the unnecessary bloat of other software.

This is probably horribly inefficient and could be solved in a number of other ways (API Keys & Hooks, etc.) but, again, this was a contrived POC type example.

# Requirements
This uses the [Selenium Python library](https://pypi.org/project/selenium/) and Chrome's [Chromedriver](http://chromedriver.chromium.org/) to achieve these results.

# Example Usage
```
OUTLOOK_EMAIL="your.email@domain.com" USER_ID="MYLOGINID" USER_PASS="MYPASSWORD" SLACK_PRIVATE_MESSAGE_URL="https://XXX.slack.com/messages/YYY" SLACK_LOGIN_URL="https://XXX.slack.com/" python calendar_refresh.py
```