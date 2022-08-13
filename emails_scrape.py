from bs4 import BeautifulSoup
import urllib.request
import re
import pandas as pd
from urllib.error import URLError, HTTPError

DataFrame = pd.read_excel('aqua.xlsx', sheetname='Sheet6')
list_of_links = DataFrame.values.tolist()
my_links = sum(list_of_links, [])
report=[]
for i in my_links:
    try:
        html_page = urllib.request.urlopen(i).read().decode('utf-8')
        soup = BeautifulSoup(html_page, 'lxml')
        regex = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+.[A-Za-z]{2,4}")
        email = re.findall(regex, html_page)
        for link in soup.findAll('a', attrs={'href': re.compile("^http://")}):
            contact_link = link.get('href')
            if 'contact' in contact_link:
                text = urllib.request.urlopen(contact_link).read().decode('utf-8')
                email_contact = re.findall(regex, text)
                report.append(i)
                report.append(email_contact)
                print(i+',', email_contact)
                break
        else:
            report.append(i)
            report.append(email)
            print(i+',', email)

    except HTTPError as e:
        report.append(i)
        report.append('empty')
        print(i+',', 'empty')
    except URLError as e:
        report.append(i)
        report.append('empty')
        print(i+',', 'empty')
    except UnicodeDecodeError as e:
        report.append(i)
        report.append('empty')
        print(i+',', 'empty')
    except ConnectionResetError as e:
        report.append(i)
        report.append('empty')
        print(i+',', 'empty')
    except ValueError as e:
        report.append(i)
        report.append('empty')
        print(i+',', 'empty')

result = [report[i:i + 2] for i in range(0, len(report), 2)]  # make list of lists from the report
writer = pd.ExcelWriter('Ylist.xlsx', engine='xlsxwriter')  # working with pandas, creating writer
df = pd.DataFrame(result, columns=['link', 'email'])  # framing data
df.to_excel(writer, sheet_name='emails', index=False)  # convert data frame to excel

#if 'Contatti' or 'Kontakt' or 'Contact' or 'contacts' or 'contact-us' in contact_link:
