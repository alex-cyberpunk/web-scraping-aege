# web-scraping-aege
## Context
  One of the tasks assigned to me at the beginning of my internship was to fill out registration forms on the energy auction website https://aege-empreendedor.epe.gov.br. This website was challenging to use as it frequently crashed, experienced downtime, and required filling in many details. My company provided the information for each project in a standard spreadsheet. I had heard from a friend at college about a method of automating clicks using Selenium in Python, and I was curious to try it out. So, I tested Selenium for web scraping on this website.

## Implementation:

  I implemented a solution that automatically fills in the website information, and it can resume from where it left off in case of interruptions. For example, if the website crashed halfway through the process, the code would continue from that point with user input. With this method, users could log into two or three different accounts and fill in the website information with more reliability and speed than manual entry. They only needed to supervise and restart the code when the website crashed.

## Possible improvements:
-Fine-tune the detection time for loading items on the page.
-Find a way to have the code open a new tab and continue from where it left off even if the page crashes, possibly saving the steps in a file.
-Implement screenshots to capture the results for later verification by team members.
