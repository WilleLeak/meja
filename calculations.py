import os
import pandas as pd


fileName = 'etsyDrawers'
dataFile = 'etsyDatabase.xlsx'

def getPrice(qty, l, w, h, unitPrice, diffBottom, bottomPrice, numPulls, pullPrice, numUMS, umPrice, numEng, engPrice, numCust, custPrice, multiplier):
    
    bottomArea = (l * w) / 144
    sideArea = (((2 * l) + (2 * w)) * h) / 144
    totalSQFT =  qty * (bottomArea + sideArea)
    totalPrice = totalSQFT * unitPrice * multiplier
    totalPrice += qty * (numPulls * pullPrice)
    totalPrice += qty * (numUMS * umPrice)
    totalPrice += qty * (numEng * engPrice)
    totalPrice += qty * (numCust * custPrice)
    
    if diffBottom == 'yes':
        totalPrice += qty * (bottomPrice)
    
    totalPrice = round(totalPrice, 2)
    
    return (totalPrice, totalSQFT)

def getCutlist(qty, l, w, h, diffBottom, numPulls, numUms, numEng, numCust, showDrawerSqft, showDrawerAddons):
    newline = '\n'
    tab = '\t'
    bottomThickness = '1/4'
    if diffBottom == 'yes':
        bottomThickness = '1/2'
        
    bottomArea = (l * w) / 144
    sideArea = (((2 * l) + (2 * w)) * h) / 144
    totalSqft = round(bottomArea + sideArea, 2)
    
   
    sqftMessage = ''
    if showDrawerSqft:
        sqftMessage = f'sqft per drawer: {totalSqft} | total drawer sqft: {totalSqft * qty}'
            
    pullMessage = ''
    engMessage = ''    
    custMessage = ''
    umMessage = ''

    if showDrawerAddons:    
        if numPulls > 0:
            pullMessage = f'pulls per box: {int(numPulls)}'
            
        if numUms > 0:
            umMessage = f'um per box: {int(numUms)}'

        if numEng > 0:
            engMessage = f'engravings per box: {int(numEng)}'
            
        if numCust > 0:
            custMessage = f'customs per box: {int(numCust)}'
        
        
    return (f'Sides: {2 * int(qty)}- {l + .5}" x {h}", {2 * int(qty)}- {w + .5}" x {h}"\n'
            f'Bottom: {int(qty)}- {l - .5}" x {w - .5}" x {bottomThickness}"\n'
            f"{tab if pullMessage != '' else ''}{pullMessage}{newline if pullMessage != '' else ''}"
            f"{tab if umMessage != '' else ''}{umMessage}{newline if umMessage != '' else ''}"
            f"{tab if engMessage != '' else ''}{engMessage}{newline if engMessage != '' else ''}"
            f"{tab if custMessage != '' else ''}{custMessage}{newline if custMessage != '' else ''}"
            f"{sqftMessage}{newline if sqftMessage != '' else ''}")
    
def saveFile(drawers, quoteNum, name, date, zipcode, shippingPrice, totalPrice, totalSqft):
    pass
    # orderDetails = {'Quote Number': quoteNum, 'Customer Name': name, 'Order Date': date, 'Zipcode': zipcode, 'Total Price': totalPrice, 'Shipping Price': shippingPrice, 'Total SqFt': totalSqft}
    # subOrderDetails = {'Quote Number': drawers}
    
    # if not os.path.isfile(dataFile):
    #     with pd.ExcelWriter(dataFile) as writer:
    #         pd.DataFrame(orderDetails).to_excel(writer, sheet_name = 'Orders', index = False)
    #         pd.DataFrame(subOrderDetails).to_excel(writer, sheet_name = 'Order Details', index = False)
            
    # else:
    #     with pd.ExcelWriter(dataFile, mode = 'a') as writer:
    #         pd.DataFrame(orderDetails).to_excel(writer, sheet_name = 'Orders', index = False, header = False)
    #         pd.DataFrame(subOrderDetails).to_excel(writer, sheet_name = 'Order Details', index = False, header = False)

    
def getDrawerWeight(drawer, materialDensities):
    sideDensity = materialDensities[drawer[0]]
    bottomDensity = materialDensities['Baltic Birch Plywood']
    qty = drawer[1]
    l = drawer[2]
    w = drawer[3]
    h = drawer[4]
    
    bottomVolume = ((l * w * .25) / 1728) * qty
    if drawer[7] == 'yes':
        bottomVolume = ((l * w * .5) / 1728) * qty
        
    sideVolume = (((2 * l) + (2 * w) * h) / 1728) * qty   
     
    bottomWeight = bottomVolume * bottomDensity
    sideWeight = sideVolume * sideDensity
    
    return round(bottomWeight + sideWeight, 2)
    
    
         
    

    