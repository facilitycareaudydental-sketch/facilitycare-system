import urllib.request
import re

url = "https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/htmlview"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')

# Find all sheet names in the html
matches = re.findall(r'\{"name":"([^"]+)"', html)
# Use a set to get unique names but keep order if possible
seen = set()
sheets = [x for x in matches if not (x in seen or seen.add(x))]
print("SHEETS:", sheets)
