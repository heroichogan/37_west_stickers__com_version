# sticker_machine.py

import win32com.client

class StickerMachine:
    _public_methods_ = [ 'excelButtonPress' ]
    _reg_progid_ = '37West.StickerMachine'
    _reg_clsid_ = '{7E26C83C-60EE-44AF-81E4-3800E357784B}' # Use "print pythoncom.CreateGuid()" to make new one (DO NOT COPY)

    def __init__( self ):  
        self._setupPublisher()

    def _setupPublisher( self ):
        self.pubApp = win32com.client.Dispatch('Publisher.Application')
        d = self.pubApp.ActiveDocument
        w = d.ActiveWindow
        w.Activate()
        w.Visible = 1

    def excelButtonPress( self, xlSheet ):
        # Make argument usable
        xlSheet = win32com.client.Dispatch( xlSheet ) #<<<<< needed?

        # Get the selection (Excel)
        data = xlSheet.Application.Selection.Value #<<<<<< must handle discontinuous selections
        
        # Get reference to (Publisher) document and the template page (Publisher)
        pubDoc = self.pubApp.Documents[0]
        templatePage = pubDoc.Pages[0]

        # From each row (Excel), create a page of stickers (Publisher)
        data = list(data)
        data.reverse()
        for row in data:
            # Create a new page (Publisher) for this row (Excel)
            page = templatePage.Duplicate()

            # Populate the stickers
            for sticker in page.Shapes: #<<<< will break if additional shapes are on screen
                # Populate text for current sticker
                sticker.GroupItems[1].TextFrame.TextRange.Text = row[0]
                sticker.GroupItems[2].TextFrame.TextRange.Text = row[1]
                sticker.GroupItems[3].TextFrame.TextRange.Text = formatAsCurrency( row[2] )

                # Force redraw to show changes
                sticker.ZOrder(2) #<<<<< msoBringForward


# Format as currency
def formatAsCurrency( xxx ):
    return '${:0.2f}'.format( xxx ) #<<<<< handle $14 case 

 
# Adapted from Hammond PyWin32
def getContiguousRange( sheet, row, col ):
    # Find the bottom row
    bottom = row
    while sheet.Cells( bottom + 1, col ).Value not in [None, '']:
        bottom = bottom + 1

    # Find the rightmost column
    right = col
    while sheet.Cells( row, right + 1 ).Value not in [None, '']:
        right = right + 1

    return sheet.Range( sheet.Cells(row, col), sheet.Cells(bottom, right)).Value



# Self-registration
if __name__ == '__main__':
    print( 'Registering COM server...' )
    import win32com.server.register
    win32com.server.register.UseCommandLine( StickerMachine, debug=1 ) # debug option
