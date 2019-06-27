# Mouser2SAPKit
To reduce the overhead created by having to enter every Mouser order item manually into the SAP system at KIT you can use the Mouser2SapKit which scrapes the website and generates the required Excel file automagically.

## Disclaimer
Fail states are not handled at all. If you are flagged as a bot user please fill in the Captcha manually. If you take too long, the script will just kill itself after the defined timeout. If the script gets stuck - kill it yourself and restart which (usually) should fix the problem.

## How to?
## Installation
´´´ bash
npm install
´´´

## Usage
´´´ bash
npm start [username] [password] [salesOrderNo]
´´´

### Example
´´´ bash
npm start myUsername mySecretPassword 123456

Starting to scrape your order ...

Username       : myUsername
Password       : mySecretPassword
Sales Order No.: 123456

Scraping done. Stored your order as 123456.xlsx!
´´´


