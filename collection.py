#! /usr/bin/env python2

"""Converting Hearthstone card data from JSON into Excel form.

This simple script reads Hearthstone cards JSON data ripped from game files and
made available at hearthstoneJSON.com into Excel spreadsheet formatted as in
Hearthstone Master Collection spreadsheet.
"""


import argparse
import json
import re
from openpyxl import Workbook

##
# Global data. This  could be embedded in the functions using them, but for clarity
# it is better to collect it here.
##

# JSON name: HMC name, HMC #Collection
SETS = {
    u'HOF'     : ("Promo",    -1),            # Hall of Fame = Promo in HMC
    u'CORE'    : ("Basic",     0),
    u'EXPERT1' : ("Classic",   1),
    u'NAXX'    : ("Naxx",      2),
    u'GVG'     : ("GvG",       3),
    u'BRM'     : ("Blackrock", 4),
    u'TGT'     : ("TGT",       5),
    u'LOE'     : ("LoE",       6),
    u'OG'      : ("TOG",       7),
    u'KARA'    : ("Kara",      8),
    u'GANGS'   : ("MSG",       9),
    u'UNGORO'  : ("Un'Goro",  10),
    u'ICECROWN' : ("KFT",     11),
    u'LOOTAPALOOZA' : ("KAC", 12)
}

# Ordering for HMC #Rarity and #Class columns
RARITY   = ["Basic", "Common", "Rare", "Epic", "Legendary"]
CLASSES  = ["", "Druid", "Hunter", "Mage", "Paladin", "Priest", "Rogue", "Shaman", "Warlock", "Warrior", "Neutral"]
# Sets currently included in standard format
STANDARD = ["TOG", "Kara", "MSG", "Un'Goro", "KFT", "KAC"]


# Add a few useful command line arguments
def handle_arguments():
    parser = argparse.ArgumentParser(description="Heartstone cards JSON to Excel conversion.")
    parser.add_argument("-i", "--input", help="input file name (JSON)", default="cards.collectible.json")
    parser.add_argument("-o", "--output", help="output file name", default="collection.xlsx")
    parser.add_argument("-s", "--sets", nargs="+", help="sets to include in output (HMC naming, default=all)")

    args = parser.parse_args()

    return args


# Read data from a JSON file (as provided by hearthstoneJSON.com) into
# a Python dictionary
def load_data(filename):
    with open(filename, "r") as fd:
        collection = json.load(fd)
    return collection


# Convert card text into ASCII, replacing special characters
# and removing markup, extra whitespace, newlines and formatting codes.
def normalize_text(ustring):
    # If a keyword is separated from text (or other keywords) by a newline, insert
    # a period, as the newlines will be removed.


    # "Wax Elemental":  <b>Taunt</b>\n<b>Divine Shield</b>
    # "Vulgar Homunculus": <b>Taunt</b>\n<b>Battlecry:</b> Deal ..
    # "Shellshifter": [x]<b>Choose One - </b>Transform\ninto a 5/3 with <b>Stealth</b>;\nor a 3/5 with <b>Taunt</b>.
    # "Al'Akir": <b>Charge, Divine Shield, Taunt, Windfury</b>
    ustring = re.sub(r'</b>\s*\n\s*', '</b>. ', ustring)

    # Replace other newlines + space with a single space.
    ustring = re.sub(r'\s*\n\s*', ' ', ustring)

    # Replace unicode characters and remove markup tags.
    replacements = [(u'\xa0', ' '), (u'\u2019', '\''),
                    ('<b>', ''), ('</b>', ''),
                    ('<i>', ''), ('</i>', ''), ('[x]', '')]
    for x,y in replacements:
        ustring = ustring.replace(x,y)

    # Effects affected by spell power are prefixed with a '$', remove it
    ustring = re.sub(r'\$(\d+)', r'\1', ustring)
    # And those restoring health by '#', do the same here
    ustring = re.sub(r'#(\d+)', r'\1', ustring)

    # Remove leading and trailing whitespace.
    ustring = ustring.strip()

    # End with a period, unless this is a single keyword. (Heuristic: if there is already a period)
    if ('.' in ustring) and (ustring[-1] not in '.)\'"'):
        ustring += '.'

    return ustring


# Card type and sub-type in HMC is a bit different from what we have in JSON,
# some of the types are in race or mechanics tags instead.
# Correct those to match existing spreadsheet conventions more closely.
def normalize_type(card):
    # By default use the type provided in JSON
    ctype   = card[u'type'].capitalize().encode('ascii')
    subtype = None

    if u'race' in card:
        subtype = card[u'race'].capitalize().encode('ascii')

    # Some types or subtypes with specific values in HMC can be identified by mechanics
    if u'mechanics' in card:
        if card[u'mechanics'] == [u'QUEST']:
            ctype = "Quest"
        if card[u'mechanics'] == [u'SECRET']:
            subtype = "Secret"
    # Death knights have no other indicator than that they are Heroes in KFT
    if ctype == "Hero" and card[u'set'] == u'ICECROWN':
        ctype = "Death Knight"
    if u'multiClassGroup' in card:
        # This needs to be revisited if multi-class mechanics ever return
        gangs = {u'GRIMY_GOONS': "Goon", u'KABAL' : "Kabal", u'JADE_LOTUS' : "Lotus"}
        subtype = gangs[card[u'multiClassGroup']]
    # Rewritings of otherwise correct values
    if subtype == "Mechanical":
        subtype = "Mech"

    return ctype, subtype

# Keywords are available in "mechanics" and "referencedTags" attributes, but not all of them
# are keywords. Also spelling is a bit different in HMC. Fix all that.
def normalize_keywords(card):
    # Keywords that will be included in "Keywords" column
    KEYWORDS = {
        u'AURA':            'Aura',
        u'BATTLECRY':       'Battlecry',
        u'CANT_ATTACK':     'Can\'t Attack',
        u'CHARGE':          'Charge',
        u'CHOOSE_ONE':      'Choose One',
        u'COMBO':           'Combo',
        u'DEATHRATTLE':     'Deathrattle',
        u'DISCOVER':        'Discover',
        u'DIVINE_SHIELD':   'Divine Shield',
        u'ENRAGED':         'Enrage',
        u'INSPIRE':         'Inspire',
        u'LIFESTEAL':       'Lifesteal',
        u'OVERLOAD':        'Overload',
        u'POISONOUS':       'Poisonous',
        u'RECRUIT':         'Recruit',
        u'SECRET':          'Secret',
        u'SPELLPOWER':      'Spell Damage',
        u'STEALTH':         'Stealth',
        u'TAUNT':           'Taunt',
        u'WINDFURY':        'Windfury'
    }

    # TODO it is questionable whether all instances should be listed as keywords.
    # In most cases only the "mechanics" should be included (so that e.g. Mad Scientist
    # and Glacial Mysteries would not have "Secret" in them.)
    # On the other hand, "Recruit" only exists in "referencedTags" and should be kept.
    # Probably the best solution is to include another field in the above dictionary
    # to indicate case by case whether keywords in referencedTags should be kept.

    # Use a set instead of list to avoid inserting an entry multiple times
    keywords = set()
    if u'mechanics' in card:
        for mechanic in card[u'mechanics']:
            if mechanic in KEYWORDS:
                keywords.add(KEYWORDS[mechanic])
    if u'referencedTags' in card:
        for tag in card[u'referencedTags']:
            if tag in KEYWORDS:
                keywords.add(KEYWORDS[tag])

    # Always return the keywords in sorted order to have some structure.
    return sorted(keywords)


# Read a single JSON entry (card = dictionary) and convert fields into another dictionary
# corresponding to conventions and column names in HMC spreadsheet.
def parse(card):
    normalized = {}

    try:
        normalized["Mana"] = card[u'cost']
        normalized["Name"] = card[u'name'].encode('ascii')
        normalized["Rarity"] = card[u'rarity'].capitalize().encode('ascii')
        if normalized["Rarity"] == "Free":
            normalized["Rarity"] = "Basic"
        normalized["Collection"] = SETS[card[u'set']][0]
        if normalized["Collection"] in STANDARD:
            normalized["Format"] = "Standard"
        else:
            normalized["Format"] = "Wild"

        normalized["Class"] = card[u'cardClass'].capitalize().encode('ascii')
        ctype, subtype = normalize_type(card)
        normalized["Type"]  = ctype
        if subtype:
            normalized["Sub-Type"] = subtype
        if u'attack' in card:
            normalized["ATK"] = card[u'attack']
        if u'health' in card:
            normalized["HP"] = card[u'health']
        if u'collectionText' in card:
            normalized["Card Text"] = normalize_text(card[u'collectionText']).encode('ascii')
        elif u'text' in card:
            normalized["Card Text"] = normalize_text(card[u'text']).encode('ascii')
        keywords = normalize_keywords(card)
        if keywords:
            # Original HMC spreadsheet is a bit inconsistent with semicolon use here.
            # Insert in between, not in the end.
            normalized["Keywords"] = "; ".join(keywords)
        normalized["#Rarity"] = RARITY.index(normalized["Rarity"])
        normalized["#Collection"] = SETS[card[u'set']][1]
        normalized["#Class"] = CLASSES.index(normalized["Class"])

    except:
        print "Error with ", card
        raise

    return normalized


# Generate a spreadsheet from reformatted card data.
def output(cards, outfile):
    wb = Workbook()
    ws = wb.active

    # Write headers
    headers = [None, "Mana", "Name", "Rarity", "Collection", "Class", "Normal", "Golden",
               "1st Copy", "2nd Copy", "Type", "Sub-Type", "ATK", "HP", "Card Text",
               "Keywords", "Format", "N-SB-Lookup", "G-SB-Lookup", "Tier List", None,
               "#Rarity", "#Collection", "#Class"]
    ws.append(headers)

    # Write data
    for card in cards:
        row = []
        for field in headers:
            if field and field in card:
                row.append(card[field])
            else:
                row.append(None)
        ws.append(row)

    wb.save(outfile)



if __name__ == "__main__":
    args = handle_arguments()
    collection = load_data(args.input)
    processed = []
    for card in collection:
        # This is a bit of a hack, but we need to exclude e.g. basic heroes (they are in collectibles list),
        # while still including Death Knights (also of Hero type).
        # For now, assume that cards without "Cost" field are not playable.
        if u'cost' not in card:
            continue
        entry = parse(card)
        # If configured, include only selected sets
        if (not args.sets) or (entry["Collection"] in args.sets):
            processed.append(entry)
    output(processed, args.output)