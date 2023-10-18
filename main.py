from functions import *

# dataset = read_json("quotes.json")
# print(get_unique_attributes_of_json(dataset, 'Author'))
# print(get_unique_attributes_of_excel(dataset, 'Author'))

dataset = read_excel("quotes_cleaner.xlsx")
quote = fetch_random_quote_excel(dataset)

print(quote)
