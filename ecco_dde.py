import os, sys, time, csv

__all__ = [
    'EccoDDE', 'DDEConnectionError', 'StateError', 'FileNotOpened',
    'WrongSession', 'ItemType', 'FolderType',
]

class DDEConnectionError(Exception):
    """Problem connecting to a DDE Server"""

class StateError(Exception):
    """Ecco is not in the expected state"""

class FileNotOpened(StateError):
    """Ecco didn't open a requested file"""

class WrongSession(StateError):
    """The expected session is not active"""

# Item Types
class ItemType(object):
    ItemText = 1
    OLELink = 2

# Folder Types
class FolderType(object):
    CheckMark  	 = 1
    Date       	 = 2
    Number     	 = 3
    Text       	 = 4
    PopUpList	 = 5










class ecco(csv.Dialect):
    delimiter = ','
    quotechar = '"'
    doublequote = True
    skipinitialspace = False
    lineterminator = ''
    quoting = csv.QUOTE_MINIMAL

csv.register_dialect("ecco", ecco)

def sz(s):
    if '\000' in s:
        return s.split('\000',1)[0]

class output(list):
    """StringIO-substitute for parsing"""
    __slots__ = ()
    write = list.append

def format(rows):
    out = output()
    fmt = csv.writer(out, 'ecco').writerows(rows)
    return '\n\r'.join(out)

def fold(seq):
    seq = iter(seq)
    return zip(seq, seq)














class EccoDDE(object):
    """A thin wrapper over the Ecco DDE API"""

    sleep = 1
    retries = 10
    filename = None
    connection = None
    server = None
    
    def __init__(self, **kw):
        cls = self.__class__
        for k, v in kw.items():
            if hasattr(cls, k):
                setattr(self, k, v)
            raise TypeError("No keyword argument "+repr(k))

    def close(self):
        """Disconnect the DDE connection and shut down the server"""
        if self.connection is not None:
            self.connection = None
        if self.server is not None:
            self.server.Shutdown()
            self.server = None

    #__del__ = close


    def assert_session(self, session_id):
        """Raise an error if active session is not `sessionId`"""
        if self.GetCurrentFile() != session_id:
            raise StateError("Attempt to close or save inactive session")

    def close_all(self):
        """Attempt to close all open files"""
        self.open()
        while True:
            try: session = self.GetCurrentFile()
            except: return
            else: self.CloseFile(session)

            
    def open(self):
        if self.connection is not None:
            return

        import win32ui, win32gui, dde, pywintypes
        if self.server is None:
            self.server = dde.CreateServer()
            self.server.Create("client")

        attempted = False
        for i in range(self.retries+1):
            try:
                win32gui.FindWindow('MauiFrame', 'Ecco')
                conn = dde.CreateConversation(self.server)
                conn.ConnectTo('Ecco', 'Ecco')
                self.connection = conn
                return
            except pywintypes.error, e:
                if e.args != (
                    2,'FindWindow','The system cannot find the file specified.'
                ):
                    raise
            except:
                t,v,tb = sys.exc_info()
                if (t,v) != ('error','ConnectTo failed'):
                    del t,v,tb; conn=None
                    raise
            if attempted:
                time.sleep(self.sleep)
            else:
                if self.filename is None:
                    import _winreg
                    self.filename = _winreg.QueryValue(
                        _winreg.HKEY_CLASSES_ROOT,
                        'NetManage EccoPro\\shell\\open\\command'
                    ).replace(' %1','')
                os.startfile(self.filename)
                attempted = True
        else:
            raise DDEConnectionError("ConnectTo failed")

    def __call__(self, cmd, *args):
        """Send `cmd` and `args` to Ecco via DDE 'request' or 'execute'

        If `args` are supplied or `cmd` is a string, a one-line csv-formatted
        request is sent.  If `cmd` is not a string, it must be an iterable of
        sequences, which will then be turned into a csv-formatted string.

        If the resulting command is more than 250 characters long, a DDE 'exec'
        will be used instead of a 'request'.  In either case, the result is        
        parsed from csv into a list of lists of strings, and then returned.
        """
        if args:
            cmd = format([(cmd,) + args])
        elif not isinstance(cmd, basestring):
            cmd = format(cmd)
        if self.connection is None:
            self.open()
        if len(cmd)>250:
            self.connection.Exec(cmd)
            data = sz(self.connection.Request('GetLastResult'))
        else:
            data = sz(self.connection.Request(cmd))
        data = data.replace('\n\r','\n').replace('\r','\n').split('\n')
        return list(csv.reader(data))

    def poke(self, cmd, *args):
        """Just like __call__(), but send a DDE poke, with no return value"""
        if args:
            cmd = format([(cmd,) + args])
        self.connection.Poke(cmd)

    def intlist(self, *cmd):
        return map(int, self(*cmd)[0])

    def one_or_many(self, cmd, ob, cvt=int):
        if hasattr(ob, '__iter__') and not isinstance(ob, basestring):
            return map(cvt, self(cmd, *ob)[0])
        else:
            return cvt(self(cmd, ob)[0][0])


    def one_or_many_to_many(self, cmd, ob, cvt=int):
        if hasattr(ob, '__iter__') and not isinstance(ob, basestring):
            return [map(cvt, row) for row in self(cmd, *ob)]
        else:
            return map(cvt, self(cmd, ob)[0])
            
    # --- "DDE Requests supported"

    def CreateFolder(self, name_or_dict, folder_type=FolderType.CheckMark):#
        """Create folders for a name or a dictionary mapping names to types
        If `name_or_dict` is a string, create a folder of `folder_type` and
        return a folder id.  Otherwise, `name_or_dict` should be a dictionary
        (or other object with an ``.items()`` method), and a dictionary mapping
        names to folder ids will be returned.
        """
        if isinstance(name_or_dict, basestring):
            return self.intlist('CreateFolder', name_or_dict, folder_type)[0]
        items = name_or_dict.items()
        items = zip(
            items,
            self.intlist('CreateFolder', [i2 for i1 in items for i2 in i1])
        )
        return dict([(k,i) for (k,t),i in items])
        
    def CreateItem(self, item, data=()):#
        """Create `item` (text) with optional data, returning new item id

        `data`, if supplied, should be a sequence of ``(folderid,value)`` pairs
        for the item to be initialized with.
        """
        return self.intlist('CreateItem', *fold(data))[0]
        
    def GetFoldersByName(self, name):
        """Return a list of folder ids for folders matching `name`"""
        return self.intlist('GetFoldersByName', name)

    def GetFoldersByType(self, folder_type=0):#
        """Return a list of folder ids whose types equal `folder_type`"""
        return self.intlist('GetFoldersByType', folder_type)


    def GetFolderItems(self, folder_id, *extra):#
        """Get the items for `folder_id`, w/optional sorting and criteria

        Examples::

            # Sort by value, descending:
            GetFolderItems(id, 'vd')

            # Sort by item text, ascending, if the folder value>26:
            GetFolderItems(id, 'ia', 'GT', 26)

            # No sort, item text contains 'foo'
            GetFolderItems(id, 'IC', 'foo')
            
        See the Ecco API documentation for the full list of supported
        operators.
        """
        return self.intlist('GetFolderItems', folder_id, *extra)

    def GetFolderName(self, folder_id):
        """Name for `folder_id` (or a list of names if id is an iterable)"""
        return self.one_or_many('GetFolderName', folder_id, str)

    def GetFolderType(self, folder_id):
        """Type for `folder_id` (or a list of types if id is an iterable)"""
        return self.one_or_many('GetFolderType', folder_id)

    #GetFolderValues
    #GetItemFolders   = singleToSeq('GetItemFolders',   'item_id(s) -> folder_ids')












    def GetItemParents(self, item_id):#
        """Return list of parent item ids (highest to lowest) of `item_id`

        If `item_id` is an iterable, return a list of lists, corresponding to
        the sequence of items.
        """
        return self.one_or_many_to_many('GetItemParents', item_id)

    def GetItemSubs(self, item_id, depth=0):#
        """itemId -> [(child_id,indent), ... ]"""
        return fold(self.intlist('GetItemSubs',depth,item_id))

    def GetItemText(self, item_id):#
        """Text for `item_id` (or a list of strings if id is an iterable)"""
        return self.one_or_many('GetItemText', item_id, str)

    def GetItemType(self, item_id):#
        """Type for `item_id` (or a list of types if id is an iterable)"""
        return self.one_or_many('GetItemType', item_id)

    def GetSelection(self):#
        """ -> [ type (1=items, 2=folders), selectedIds]"""
        res = [ map(int,line) for line in self('GetSelection') ]
        res[0] = res[0][0]
        return res

    def GetVersion(self):
        """Return the Ecco API protocol version triple (major, minor, rev#)"""
        return self.intlist('GetVersion')
        











    def NewFile(self):
        """Create a new 'Untitled' file, returning a session id"""
        return int(self('NewFile')[0][0])

    def OpenFile(self, pathname):#
        """Open or switch to `pathname` and return a session ID

        If the named file was not actually opened (not found, corrupt, etc.),
        a ``ecco_dde.FileNotOpened`` error will be raised instead.
        """
        result = self(format([['OpenFile', pathname]]))[0]
        result = result and int(result[0]) or 0
        if not result:
            raise FileNotOpened(pathname)
        return result

    #PasteOLEItem Flags, [ ItemID ], ( FolderID, FolderValue ) * -> ItemID

    # --- "Extended DDE Requests"

    #GetChanges

    def GetViews(self):
        """Return a list of the view ids of all views in current session"""
        return self.intlist('GetViews')

    def GetViewNames(self, view_id=None):
        """Return one or more view names

        If `view_id` is an iterable, this returns a list of view names for the
        corresponding view ids.  If `view_id` is ``None`` or not given, this
        returns a list of ``(name, id)`` pairs for all views in the current
        session.  Otherwise, the name of the specified view is returned.
        """
        if view_id is None:
            views = self.intlist('GetViews')
            return zip(self.GetViewNames(views), views)
        else:
            return self.one_or_many('GetViewNames', view_id, str)


    def GetViewFolders(self, view_id):
        """Folder ids for `view_id` (or list of lists if id is an iterable)"""
        return self.one_or_many_to_many('GetViewFolders', view_id)

    def GetPopupValues(self, folder_id):#
        """Popup values for `folder_id` (or list of lists if id is iterable)"""
        return self.one_or_many_to_many('GetPopupValues', folder_id, str)
       
    def GetFolderOutline(self):
        """Return a list of ``(folderid, depth)`` pairs for the current file"""
        return fold(self.intlist("GetFolderOutline"))

    def GetViewColumns(self, view_id):#
        """Folderids for `view_id` cols (or list of lists if id is iterable)"""
        return self.one_or_many_to_many('GetViewColumns', view_id)

    def GetViewTLIs(self, view_id):#
        """Return a list of ``(folder_id, itemlist)`` pairs for `view_id`"""
        rows = self('GetViewTLIs', view_id)
        for pos, row in enumerate(rows):
            rows[pos] = row.pop(0), row
        return rows

    def GetOpenFiles(self):
        """Return a list of session IDs for all currently-open files"""
        return self.intlist("GetOpenFiles")

    def CreateView(self, name, folder_ids):
        assert folder_ids, "Must include at least one folder ID!"
        return self.intlist('CreateView', name, *folder_ids)[0]

    def GetFolderAutoAssignRules(self, folder_id):#
        """Get list of strings defining auto-assign rules for `folder_id`"""
        return self('GetFolderAutoAssignRules', folder_id)[0]

    def GetCurrentFile(self):
        """Return the session id of the active file"""
        return int(self('GetCurrentFile')[0][0])



    def GetFileName(self, session_id):
        """Return the file name for the given session ID"""
        return self(format([['GetFileName', session_id]]))[0][0]

    # --- "DDE Pokes supported"

    def ChangeFile(self, session_id):#
        """Switch to the designated `session_id`"""
        # Alas, the poke doesn't always work, at least not in my Ecco...
        if self.GetCurrentFile()!=session_id:
            self.poke('ChangeFile', session_id)
            # So we may have to use OpenFile instead:
            if self.GetCurrentFile()!=session_id:
                self.OpenFile(self.GetFileName(session_id))

    def CloseFile(self, session_id):
        """Close the designated session, *without* saving it"""
        self.assert_session(session_id)
        self.poke('CloseFile')

    #CopyOLEItem ItemID
    #InsertItem ItemID, flags, ItemID *

    def RemoveItem(self, item_id):#
        """Delete `item_id` (can be an iterable of ids)"""
        if hasattr(item_id, '__iter__'):
            self.poke('RemoveItem', *item_id)
        else:
            self.poke('RemoveItem', item_id)

    def SaveFile(self, session_id, pathname=None):#
        """Save the designated session to `pathname`; fails if not current"""
        self.assert_session(session_id)
        if pathname:
            self.poke('SaveFile', pathname)
        else:
            self.poke('SaveFile')




    def SetFolderName(self, folder_id, name):#
        """Set the name of `folder_id` to `name`"""
        self.poke('SetFolderName', folder_id, name)

    #SetFolderValues < FolderID * > < ItemID * > < FolderValue * > *
    #SetItemText ( ItemID, "text" ) *
    #ShowPhoneBookItem,<itemId>[,<bClear>]
    
    # --- "Extended DDE Pokes"

    def ChangeView(self, view_id):#
        """Display the specified view"""
        self.poke('ChangeView', view_id)

    def AddCompView(self, view_id):#
        """Add `view_id` as a composite view to the current view"""
        self.poke('AddCompView', view_id)
        
    def RemoveCompView(self, view_id):#
        """Remove the specified view from the current view's composite views"""
        self.poke('RemoveCompView', view_id)

    #SetCalDate Date

    def DeleteView(self, view_id):#
        """Delete the specified view"""
        self.poke('DeleteView', view_id)

    #AddFileToMenu FilePath IconID
    #AddColumnToView ViewID FolderID*
    #AddFolderToView ViewID FolderID*










def additional_tests():
    import doctest
    return doctest.DocFileSuite(
        'README.txt',
        optionflags=doctest.ELLIPSIS|doctest.NORMALIZE_WHITESPACE,
    )



































