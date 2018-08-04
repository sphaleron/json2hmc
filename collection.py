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
    u'LOOTAPALOOZA' : ("KnC", 12),
    u'GILNEAS' : ("Woods", 13),
    u'BOOMSDAY' : ("Boom", 14)
}

# Ordering for HMC #Rarity and #Class columns
RARITY   = ["Basic", "Common", "Rare", "Epic", "Legendary"]
CLASSES  = ["", "Druid", "Hunter", "Mage", "Paladin", "Priest", "Rogue", "Shaman", "Warlock", "Warrior", "Neutral"]
# Sets currently included in standard format
STANDARD = ["Un'Goro", "KFT", "KnC", "Woods", "Boom"]


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
    # If a keyword is separated from other content by a newline, insert
    # a period. Newlines will be removed.

    # "Wax Elemental":  <b>Taunt</b>\n<b>Divine Shield</b>
    # "Vulgar Homunculus": <b>Taunt</b>\n<b>Battlecry:</b> Deal ..
    # "Shellshifter": [x]<b>Choose One - </b>Transform\ninto a 5/3 with <b>Stealth</b>;\nor a 3/5 with <b>Taunt</b>.
    # "Al'Akir": <b>Charge, Divine Shield, Taunt, Windfury</b>
    # "Dragonhawk Rider": <b>Inspire:</b> Gain <b>Windfury</b>\nthis turn.
    # "Skycap'n Kragg": <b>Charrrrrge</b>\nCosts (1) less for each friendly Pirate.
    ustring = re.sub(r'</b>\s*\n\s*([^a-z])', r'</b>. \1', ustring)

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
    ctype   = card[u'type'].capitalize()
    subtype = None

    if u'race' in card:
        subtype = card[u'race'].capitalize()

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
        # Do not overwrite existing subtype, e.g. Elemental for Jade Spirit
        if subtype == None:
            subtype = gangs[card[u'multiClassGroup']]
    # Rewritings of otherwise correct values
    if subtype == "Mechanical":
        subtype = "Mech"

    return ctype, subtype

# Keywords are available in "mechanics" and "referencedTags" attributes, but not all of them
# are keywords. Also spelling is a bit different in HMC. Fix all that.
def normalize_keywords(card):
    # Keywords that will be included in "Keywords" column, and a flag if also referencedTags
    # appearances are considered.
    KEYWORDS = {
        u'ADAPT':           ('Adapt', True),
        u'AURA':            ('Aura', False),
        u'BATTLECRY':       ('Battlecry', False),
        u'CANT_ATTACK':     ('Can\'t Attack', False),
        u'CHARGE':          ('Charge', False),
        u'CHOOSE_ONE':      ('Choose One', False),
        u'COMBO':           ('Combo', False),
        u'DEATHRATTLE':     ('Deathrattle', False),
        u'DISCOVER':        ('Discover', True),
        u'DIVINE_SHIELD':   ('Divine Shield', False),
        u'ECHO':            ('Echo', False),
        u'FREEZE':          ('Freeze', True),
        u'INSPIRE':         ('Inspire', False),
        u'LIFESTEAL':       ('Lifesteal', False),
        u'MODULAR':         ('Magnetic', False),
        u'OVERLOAD':        ('Overload', False),
        u'POISONOUS':       ('Poisonous', False),
        u'RECRUIT':         ('Recruit', True),
        u'RUSH':            ('Rush', False),
        u'SECRET':          ('Secret', False),
        u'SPELLPOWER':      ('Spell Damage', True),
        u'STEALTH':         ('Stealth', False),
        u'TAUNT':           ('Taunt', True),
        u'WINDFURY':        ('Windfury', False)
    }

    # In most cases only the keywords in field "mechanics" should be included (so that
    # e.g. Mad Scientist and Glacial Mysteries would not have "Secret" in them.)
    # On the other hand, "Recruit" only exists in "referencedTags" and should be kept.
    # HMC sheet is somewhat inconsistent in it's use of keywords, the above flags are
    # chosen to best fit the previous versions with manual input.

    # Using a set instead of list to avoid inserting an entry multiple times
    keywords = set()
    if u'mechanics' in card:
        for mechanic in card[u'mechanics']:
            if mechanic in KEYWORDS:
                keywords.add(KEYWORDS[mechanic][0])
    if u'referencedTags' in card:
        for tag in card[u'referencedTags']:
            if tag in KEYWORDS and KEYWORDS[tag][1]:
                keywords.add(KEYWORDS[tag][0])

    # Always return the keywords in sorted order to have some structure.
    return sorted(keywords)


# Read a single JSON entry (card = dictionary) and convert fields into another dictionary
# corresponding to conventions and column names in HMC spreadsheet.
def parse(card):
    normalized = {}

    try:
        normalized["Mana"] = card[u'cost']
        normalized["Name"] = card[u'name']
        normalized["Rarity"] = card[u'rarity'].capitalize()
        if normalized["Rarity"] == "Free":
            normalized["Rarity"] = "Basic"
        if u'set' not in card:
            # Workaround for cards HOF'ed in Witchwood missing set attribute
            card[u'set'] = u'HOF'
        normalized["Collection"] = SETS[card[u'set']][0]
        if normalized["Collection"] in STANDARD:
            normalized["Format"] = "Standard"
        else:
            normalized["Format"] = "Wild"

        normalized["Class"] = card[u'cardClass'].capitalize()
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

    # For really old versions of openpyxl
    ws = wb.get_active_sheet()
    #ws = wb.active

    # Write headers
    headers = [None, "Mana", "Name", "Rarity", "Collection", "Class", "N", "G",
               "W1", "W2", "Type", "Sub-Type", "ATK", "HP", "Card Text",
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
