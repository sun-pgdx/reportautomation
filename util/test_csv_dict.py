import csv
import pprint
pp = pprint.PrettyPrinter(indent=4)

file = 'trigger.csv'
with open(file, mode='r') as infile:
  reader = csv.reader(infile)
#  mydict = {rows[0]:rows[1] for rows in reader}
  mydict = dict((rows[0],rows[1]) for rows in reader)

  pp.pprint(mydict)
