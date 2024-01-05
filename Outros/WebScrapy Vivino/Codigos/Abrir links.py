import webbrowser

filename = "F:\COMPRAS\Automações\Automação - Vivino\Outros\Links.txt"

with open(filename, "r") as file:
    links = file.readlines()

links = [link.strip() for link in links]

for link in links:
    webbrowser.open(link)
