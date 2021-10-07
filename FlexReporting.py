run = input("Which report do you want to run? (Through), (Scrap), (Both) ")
from FlexThroughput import flex_through
from FlexScrap import flex_scrap

if run == str.lower("Through"):
    flex_through()
elif run == str.lower("Scrap"):
    flex_scrap()
else:
    flex_through()
    flex_scrap()