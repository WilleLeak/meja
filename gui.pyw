import json
import os
from xml.dom import pulldom
import PySimpleGUI as sg
import calculations as calc
from datetime import datetime, timedelta

# ======================GLOBAL VARIABLES============================

# window size
windowSizeX = 1830
windowSizeY = 730

# fonts
headerFont = ('Arial Bold', 15)
textFont = ('Arial Bold', 12)
errorFont = ('Arial Bold', 20)
smallFont = ('Arial Bold', 10)

# current date and expiration date
date = datetime.now()
expirationDate = date + timedelta(days=14)

# dropdown options
materials = [] # filled out in settings
yay = ['yes', 'no']
prices = [6 + i * .25 for i in range(int((12 - 6) / .25) + 1)]


# misc
tablePadding = (1, 1)
wordPadding = (0, 0)
smallInput = (8, 1)
textSize = (7, 1)
dropdownSize = (6, 1)
options = ['material', 'qty', 'length', 'width', 'height', 'unitPrice', 'bott', 'bottPrice', 'numPulls', 'pullPrice', 'numUMs', 'umPrice', 'numEng', 'engPrice', 'numCustom', 'custDesc', 'custPrice']
settingsFile = 'settings'
defaultSettings = [
    {'materialMultipliers': {
    'Baltic Birch Plywood': 1,
    'Maple': 1,
    'Birch': 1,
    'Alder': 1,
    'Black Walnut': 3}
    },
    {'shippingInfo' : {
        'deliveryTime': '1-2 weeks',
        'turnaround': '3-4 days'}
    },
    {'messageSettings': {
        'showIndividualDrawerPrices': False,
        'showUpgradeMessage': True}
    },
    {'materialDensities': {
        'Baltic Birch Plywood': 36.5,
        'Maple': 34.5,
        'Birch': 33.5,
        'Alder': 31,
        'Black Walnut': 32.5}
    },
    {'cutlistSettings': {
        'showTotalWeight': True,
        'showTotalSqft': True,
        'showDrawerSqft': True} 
    }
]

# =================END OF GLOBAL VARIABLES==========================

# ============================SETTINGS==============================

# access settings
    # order as follows:
        # index 0 = materials/multipliers
        # index 1 = shipping info
        # index 2 = output message settings
        # index 3 = material densities (lbs/ft^3)
        # index 4 = cutlist settings

def readSettings():
    global materials
    global materialMultipliers
    global deliveryTime
    global turnaround
    global showIndividualDrawerPrices
    global showUpgradeMessage
    global materialDensities
    global showTotalWeight
    global showTotalSqft
    global showDrawerSqft
    global showDrawerAddons
    
    existingSettings = []
    if os.path.exists(settingsFile):
        with open(settingsFile, 'r') as file:
            existingSettings = json.load(file)
            
    else:
        existingSettings = [
            {'materialMultipliers': {
            'Baltic Birch Plywood': 1,
            'Maple': 1,
            'Birch': 1,
            'Alder': 1,
            'Black Walnut': 3}
            },
            {'shippingInfo' : {
                'deliveryTime': '1-2 weeks',
                'turnaround': '3-4 days'}
            },
            {'messageSettings': {
                'showIndividualDrawerPrices': False,
                'showUpgradeMessage': True}
            },
            {'materialDensities': {
                'Baltic Birch Plywood': 34.5,
                'Maple': 32.5,
                'Birch': 31.5,
                'Alder': 30,
                'Black Walnut': 30.5}
            },
            {
                'cutlistSettings': {
                    'showTotalWeight': True,
                    'showTotalSqft': True,
                    'showDrawerSqft': True,
                    'showDrawerAddons': True
                }
            }
        ]

        with open(settingsFile, 'w') as file:
            json.dump(existingSettings, file, indent = 4)
            
    materialMultipliers = existingSettings[0]['materialMultipliers']
    materials = list(existingSettings[0]['materialMultipliers'].keys())
    deliveryTime = existingSettings[1]['shippingInfo']['deliveryTime']
    turnaround = existingSettings[1]['shippingInfo']['turnaround']
    showIndividualDrawerPrices = existingSettings[2]['messageSettings']['showIndividualDrawerPrices']
    showUpgradeMessage = existingSettings[2]['messageSettings']['showUpgradeMessage']
    materialDensities = existingSettings[3]['materialDensities']
    showTotalWeight = existingSettings[4]['cutlistSettings']['showTotalWeight']
    showTotalSqft = existingSettings[4]['cutlistSettings']['showTotalSqft']
    showDrawerSqft = existingSettings[4]['cutlistSettings']['showDrawerSqft']
    showDrawerAddons = showDrawerSqft = existingSettings[4]['cutlistSettings']['showDrawerAddons']
    
     
readSettings()
 
# ============================END OF SETTINGS=======================

# ================HELPER FUNCTIONS==================================

def addDrawer(rowCounter):
    row = [
        sg.pin(
            sg.Col([
                [
                sg.Combo(materials, key = ('material', rowCounter), size = (20, 1), font = smallFont, pad = (1, 0), default_value = 'Baltic Birch Plywood'),
                sg.In(key = ('qty', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('length', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('width', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('height', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('unitPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding, default_text = '10'),
                sg.Combo(yay, key = ('bott', rowCounter), font = smallFont, size = (6, 1), pad = tablePadding, default_value = 'no'),
                sg.In(key = ('bottPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding, default_text = '6'),
                sg.In(key = ('numPulls', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('pullPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding, default_text = '6'),
                sg.In(key = ('numUMs', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('umPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding, default_text = '6'),
                sg.In(key = ('numEng', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('engPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding, default_text = '8'),
                sg.In(key = ('numCustom', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.In(key = ('custDesc', rowCounter), size = (30, 1), font = smallFont, pad = tablePadding),
                sg.In(key = ('custPrice', rowCounter), size = smallInput, font = smallFont, pad = tablePadding),
                sg.Button('X', border_width = 0, key = ('del', rowCounter))
                ]    
            ],
                key = ('row', rowCounter)
            )
        )
    ]
    return row

def createHeading(headerName, size):
    row = sg.Text(headerName, pad = tablePadding, size = size, justification = 'center', font = smallFont)
    return row

def createSingleMessage(drawer):
    tab = '\t'
    newLine = '\n'
    widthMessage = ''
    if drawer[6] == 'no' and (drawer[2] >= 24 or drawer[3] >= 24):
        widthMessage = f'I recommend a 1/2 inch bottom since the drawer is very long/wide. This would add ${round(drawer[1] * drawer[7], 2)} to your order.'
        
    individualPriceMessage = ''   
    if showIndividualDrawerPrices:
        individualPriceMessage = f"Total drawer price ${drawer[17]}\n"
        
    pullMessage = ''    
    if drawer[8] > 0:
        pullMessage = f'{int(drawer[8])} pulls per box'
    
    umMessage = ''    
    if drawer[10] > 0:
        umMessage = f'{int(drawer[10])} undermounts per box'
        
    engMessage = ''
    if drawer[12] > 0:
        engMessage = f'{int(drawer[12])} engravings per box'
        
    custMessage = ''
    custDetails = f'({drawer[15]})'
    if drawer[14] > 0:
        custMessage = f"{int(drawer[14])} custom addons per box {custDetails if drawer[15] != '' else ''}"
        
        
    return (f"({int(drawer[1])}) {'box' if drawer[1] == 1 else 'boxes'} {drawer[2]}\" x {drawer[3]}\" x {drawer[4]}\" (W = Width | L = Length | H = Height)\n" 
            f"Material | {drawer[0]} | {'1/4 inch' if drawer[6] == 'no' else '1/2 inch'} Birch plywood bottom included\n" 
            f"{widthMessage}{newLine if widthMessage != '' else ''}"
            f"{'Extra Additions:' if pullMessage != '' or umMessage != '' or engMessage != '' or custMessage != '' else ''}{newLine if pullMessage != '' or umMessage != '' or engMessage != '' or custMessage != '' else ''}"
            f"{tab if pullMessage != '' else ''}{pullMessage}{newLine if pullMessage != '' else ''}"
            f"{tab if umMessage != '' else ''}{umMessage}{newLine if umMessage != '' else ''}"
            f"{tab if engMessage != '' else ''}{engMessage}{newLine if engMessage != '' else ''}"
            f"{tab if custMessage != '' else ''}{custMessage}{newLine if custMessage != '' else ''}"
            f"{individualPriceMessage}\n")

def editSettings(filepath):
    existingSettings = []
    with open(filepath, 'r') as file:
        existingSettings = json.load(file)
    
    prettySettings = json.dumps(existingSettings, indent = 4)
    editorLayout = [
        [sg.Text('Edit Settings', font = textFont)],
        [sg.Text('Punctuation, spelling and format are VERY important.\nDeleting the settings file and reopening the program will restore defaults', font = smallFont)],
        [sg.Multiline(prettySettings, key = 'settings', size = (64, 45))],
        [sg.Button('Save', key = 'save'), sg.Push(), sg.Button('Restore Defaults', key = 'restore')]
    ]
    
    window = sg.Window('Settings Editor', editorLayout, finalize = True)
    
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        
        if event == 'save':
            newSettings = values['settings']
            settingsJSON = json.loads(newSettings)
            with open(filepath, 'w') as file:
                json.dump(settingsJSON, file, indent = 4)
            break
        
        if event == 'restore':
            with open(settingsFile, 'w') as file:
                json.dump(defaultSettings, file, indent = 4)
            break
        
    window.close()

def getDrawers():
    activeRows = []
    for i in range(rowCounter + 1):
        if window[('row', i)].visible:
            drawer = []
            for option in options:
                if option != 'custDesc' and values[(option, i)] == '':
                    drawer.append(0)
                elif option != 'custDesc' and option != 'material' and option != 'bott':
                    drawer.append(float(values[(option, i)]))
                else:
                    drawer.append(values[option, i])
            activeRows.append(drawer)
            
    for val in activeRows:
        cost = calc.getPrice(qty = val[1], l = val[2], w = val[3], h = val[4], 
                    unitPrice = val[5], 
                    numPulls = val[8], pullPrice = val[9],
                    numUMS = val[10], umPrice = val[11],
                    numEng = val[12], engPrice = val[13],
                    numCust = val[14], custPrice = val[16],
                    diffBottom = val[6], bottomPrice = val[7],
                    multiplier = materialMultipliers[val[0]])
        val.append(cost[0])
        val.append(cost[1])
    return activeRows

# ================END OF HELPER FUNCTIONS===========================

# ================GUI LAYOUT========================================
sg.theme('DarkGrey')

quoteLayout = [
    [sg.Text('Quote Number', font = textFont), sg.InputText(key = 'quoteNum', size = (10,1), font = textFont)],
    [sg.Text('Client Name', font = textFont), sg.InputText(key = 'clientName', size = (30, 1), font = textFont)],
    [sg.Text('Today\'s Date:', font = textFont), sg.Text(date.strftime('%B %d, %Y'), font = textFont), sg.Text('Expiration date:', font = textFont), sg.Text(expirationDate.strftime('%B %d, %Y'), font = textFont)],
    [sg.Text('Zipcode:', font = textFont), sg.InputText(key = 'zipcode', size = smallInput, font = textFont), sg.Text('Shipping Price:', font = textFont), sg.InputText(key = 'shipping', size = smallInput, font = textFont),]  
]

optionsLayout = [
    [
        createHeading('Material', (20, 1)),
        createHeading('Qty', textSize),
        createHeading('Length', textSize),
        createHeading('Width', textSize),
        createHeading('Height', textSize),
        createHeading('Price', textSize),
        createHeading('1/2 Bot', textSize),
        createHeading('Price', textSize),
        createHeading('# Pulls', textSize),
        createHeading('Price', textSize),
        createHeading('# UMs', textSize),
        createHeading('Price', textSize),
        createHeading('# Eng', textSize),
        createHeading('Price', textSize),
        createHeading('# Cstm', textSize),
        createHeading('Description', (26, 1)),
        createHeading('Price', textSize)
    ]
]

outputLayout = [
    sg.Col([
        [sg.Multiline(key = 'quickInfo', size = (20, 1))]
    ])
]

# sg.vtop(outputLayout)
materialLayout = [
    [sg.vtop(sg.Col([addDrawer(0)], key = 'drawerRow')), sg.Push(), sg.vtop(sg.Multiline(key = 'quickInfo', size = (60, 30), no_scrollbar = True))]
]

mainLayout = [
    [sg.Frame('', quoteLayout), sg.Push(), sg.vtop(sg.Multiline(key = 'cutlist', size = (58, 10), default_text = 'CUTLIST'))], 
    optionsLayout,
    materialLayout,
    [sg.vtop(sg.Button('Add Drawer', key = 'addDrawer')), 
     sg.vtop(sg.Button('Submit', key = 'submit')),
     # sg.vtop(sg.Button('Clear All', key = 'clearAll')),
     sg.Push(), 
     sg.In(key = 'settingsFilePath', visible = False, enable_events = True), 
     sg.FileBrowse('Edit Settings'), sg.vtop(sg.Button('Save to File', key = 'saveFile'))]
]

# =================END OF GUI LAYOUT================================

# =================EVENT LOOP=======================================

window = sg.Window('Etsy Drawer Pricing', mainLayout, size = (windowSizeX, windowSizeY))

rowCounter = 0
drawers = []
totalPrice = 0
totalSqft = 0
totalDrawers = 0
upgradePrice = 0
outputMessage = ''

while True:
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    
    if event == 'addDrawer':
        rowCounter += 1
        window.extend_layout(window['drawerRow'], [addDrawer(rowCounter)])
    elif event[0] == 'del':
        window[('row', event[1])].update(visible=False)
        
    if event == 'submit':
        clientName = values['clientName']
        outputMessage = f'Hi {clientName}, thank you for the inquiry.\nPlease let me know if anything is incorrect.\n\n'
        totalPrice = 0
        totalSqft = 0
        drawers = getDrawers()
        
        totalWeight = 0
        upgradePrice = 0
        upgradeMessage = ''
        for val in drawers:
            totalPrice += val[17]
            totalSqft += val[18]
            totalDrawers += val[1]
            outputMessage += createSingleMessage(val)
            totalWeight += calc.getDrawerWeight(val, materialDensities)
    
            if (val[2] >= 24 or val[3] >= 24) and showUpgradeMessage:
                upgradePrice += val[7] * val[1]
                
        shippingPrice = values['shipping']     
        if shippingPrice != '':
            shippingPrice = float(values['shipping'])
        else:
            shippingPrice = 0
                
        if showUpgradeMessage and upgradePrice > 0:
            updatedTotal = round(totalPrice + upgradePrice, 2)
            upgradeMessage = f'Upgrade to 1/2 inch bottoms for ${upgradePrice}. The new total price would be ${updatedTotal + shippingPrice}\n'
            
        outputMessage = outputMessage + f'Drawer Cost ${round(totalPrice, 2)}\n'    
        outputMessage = outputMessage + f'Shipping ${round(shippingPrice, 2)}\n'
        outputMessage = outputMessage + f'***Total Costs ${round(totalPrice + shippingPrice, 2)}***\n'
        outputMessage = outputMessage + f'{upgradeMessage}'
        outputMessage = outputMessage + f'The shipping is a bit high. Before shipping I will look for a better rate and refund any difference.\n'
        outputMessage = outputMessage + f'Prices are good for 14 days\n'
        outputMessage = outputMessage + f'{deliveryTime} for delivery (Currently run at {turnaround} turnaround)\n'
        
            
        globalCutlist = 'CUTLIST\n'
        count = 1
        for val in drawers:
            cutlist = f'{val[0]} drawer set {count}\n'
            cutlist = cutlist + calc.getCutlist(qty = val[1], l = val[2], w = val[3], h = val[4], diffBottom = val[6], numPulls = val[8], numUms = val[10], numEng = val[12], numCust = val[14],
                                                showDrawerSqft = showDrawerSqft, showDrawerAddons = showDrawerAddons)
            cutlist = cutlist + '\n'
            globalCutlist = globalCutlist + cutlist
            count += 1
            
        if showTotalSqft:
            globalCutlist = globalCutlist + f'grand total sqft: {round(totalSqft, 2)}\n'
            
        if showTotalWeight:
            globalCutlist = globalCutlist + f'total weight: {round(totalWeight, 2)} lbs'
            
        window['cutlist'].update(globalCutlist)
        window['quickInfo'].update(outputMessage)
        
    if event == 'saveFile':
        totalPrice = 0
        totalSqft = 0
        drawers = getDrawers()

        for val in drawers:
            totalPrice += val[17]
            totalSqft += val[18]
        
        quoteNum = 0
        if values['quoteNum'] != '':
            quoteNum = int(values['quoteNum'])
            
        shippingPrice = 0
        if type(values['shipping']) == int and values['shipping'] > 0:
            shippingPrice = values['shipping']
        
        calc.saveFile(drawers = drawers, quoteNum = quoteNum, name = values['clientName'], date = date.strftime('%B %d, %Y'), zipcode = values['zipcode'], shippingPrice = shippingPrice, totalSqft = totalSqft, totalPrice = totalPrice)
        
    if event == 'settingsFilePath':
        filepath = values['settingsFilePath']
            
        editSettings(filepath)
        readSettings()
        
    if event == 'clearAll':
        pass
        
    
window.close()
 
# ==============END OF EVENT LOOP===================================