import tkinter

class customFlatButton(tkinter.Button):
    ''' - Class for custom-style button
        - Meant to be used in tkinter frame, labelframe
        - and everywhere tkinter.Button can be used
        - Author(s): @author1 (See authors.txt file)
    '''
    def __init__(self, master, text, backgroundColor, buttonTextColor = 'black', accentColor = 'Whith', command = None):
        super().__init__(
            master = master,
            text = text,
            background = backgroundColor,
            foreground = buttonTextColor,
            activebackground = accentColor,
            relief = tkinter.FLAT,
            command = command
        )
        self.accentColor = accentColor,
        self.backgroundColor = backgroundColor
        self.bind('<Enter>', self._onEnter)
        self.bind('<Leave>', self._onLeave)

    def _onEnter(self, event):
        self.configure(background = self.accentColor)
    
    def _onLeave(self, event):
        self.configure(background = self.backgroundColor)
    

