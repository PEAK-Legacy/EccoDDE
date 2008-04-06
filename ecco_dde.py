import os, sys, time, csv, datetime

__all__ = [
    'EccoDDE', 'DDEConnectionError', 'StateError', 'FileNotOpened',
    'WrongSession', 'ItemType', 'FolderType', 'InsertLevel', 'OLEMode',
    'SelectionType', 'format_date', 'format_datetime',
]

class DDEConnectionError(Exception):
    """Problem connecting to a DDE Server"""

class StateError(Exception):
    """Ecco is not in the expected state"""

class FileNotOpened(StateError):
    """Ecco didn't open a requested file"""

class WrongSession(StateError):
    """The expected session is not active"""


def format_date(dt):
    if hasattr(dt, 'strftime'):
        return dt.strftime("%Y%m%d")
    return dt

def format_datetime(dt):
    if hasattr(dt, 'strftime'):
        return dt.strftime("%Y%m%d%H%M%s")
    return dt


def additional_tests():
    import doctest
    return doctest.DocFileSuite(
        'README.txt',
        optionflags=doctest.ELLIPSIS|doctest.NORMALIZE_WHITESPACE,
    )



# Item Types
class ItemType(object):
    ItemText = 1
    OLE = 2

# Selection Types
class SelectionType(object):
    Items = 1
    Folders = 2
    Nothing = 0

# Folder Types
class FolderType(object):
    CheckMark  	 = 1
    Date       	 = 2
    Number     	 = 3
    Text       	 = 4
    PopUpList	 = 5

# Insert locations
class InsertLevel(object):
    Indent  = 'd'   # "first daughter"
    Outdent = 'a'   # "next aunt"
    Same    = 's'   # "next sister"

# OLE Modes
class OLEMode(object):
    Link = 1
    Embed = 2












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
    """StringIO-substitute for csv writing"""
    __slots__ = ()
    write = list.append

def format(rows):
    out = output()
    fmt = csv.writer(out, 'ecco').writerows(rows)
    return '\n\r'.join(out)

def fold(seq):
    seq = iter(seq)
    return zip(seq, seq)

def unfold(seq):
    return [i2 for i1 in seq for i2 in i1]











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
            session = self.GetCurrentFile()
            if session is None: return
            self.CloseFile(session)


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
        data = args and format([args]) or ''
        if self.connection is None:
            self.open()
        self.connection.Poke(cmd, data)

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

    def CreateFolder(self, name_or_dict, folder_type=FolderType.CheckMark):
        """Create folders for a name or a dictionary mapping names to types
        If `name_or_dict` is a string, create a folder of `folder_type` and
        return a folder id.  Otherwise, `name_or_dict` should be a dictionary
        (or other object with an ``.items()`` method), and a dictionary mapping
        names to folder ids will be returned.
        """
        if isinstance(name_or_dict, basestring):
            return self.intlist('CreateFolder', folder_type, name_or_dict)[0]
        items = name_or_dict.items()
        items = zip(
            items,
            self.intlist('CreateFolder', *unfold([(j,i) for i,j in items]))
        )
        return dict([(k,i) for (k,t),i in items])

    def CreateItem(self, item, data=()):
        """Create `item` (text) with optional data, returning new item id

        `data`, if supplied, should be a sequence of ``(folderid,value)`` pairs
        for the item to be initialized with.
        """
        return self.intlist('CreateItem', item, *unfold(data))[0]

    def GetFoldersByName(self, name):
        """Return a list of folder ids for folders matching `name`"""
        return self.intlist('GetFoldersByName', name)

    def GetFoldersByType(self, folder_type=0):
        """Return a list of folder ids whose types equal `folder_type`"""
        return self.intlist('GetFoldersByType', folder_type)


    def GetFolderItems(self, folder_id, *extra):
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















    def GetFolderValues(self, item_ids, folder_ids):
        """Return folder values for specified folders and items

        `item_ids` can be a single item ID, or a sequence.  If it's a sequence,
        the return value is a sequence ordered by the input item ids.

        `folder_ids` can be a single folder ID, or a sequence.  If it's a
        sequence, the data returned for each item is a sequence ordered by the
        input folder ids.

        If a single item ID and single folder ID are used, that folder value is
        returned for that item.  If a single folder ID is used and multiple
        item IDs, the result is a list of values, one value per item.  If
        multiple folder IDs are used, and only one item ID, then the result is
        a single list containing the values for that one item.

        In other words, depending on the input, you'll get either a value, a
        list of values (either different folders for one item, or one folder
        for different items), or a list of lists of values.
        """
        cmd = [['GetFolderValues'], []]
        for inp,out in zip([item_ids,folder_ids], cmd):
            if hasattr(inp,'__iter__'):
                out.extend(inp)
            else:
                out.append(inp)
        data = self(cmd)
        if not hasattr(folder_ids, '__iter__'):
            for i,v in enumerate(data): data[i], = v or ('',)
        if not hasattr(item_ids, '__iter__'):
            data, = data
        return data









    def GetItemFolders(self, item_ids):
        """Get the folders for `item_ids`

        If `item_ids` is iterable, each element must be either an item id or
        an iterable of item ids, representing a set of items for whom the
        folders should be retrieved.  The result is a list of lists of folders,
        corresponding to the order of the input iterable.

        If `item_ids` is not iterable, the return value is a list of folder ids
        containing that one specific item.
        """
        if hasattr(item_ids, '__iter__'):
            data = []
            for item in item_ids:
                if hasattr(item, '__iter__'):
                    data.append(list(item))
                else:
                    data.append([item])
            data[0].insert(0,'GetItemFolders')
            return [map(int,d) for d in self(data)]
        else:
            return self.intlist('GetItemFolders',item_ids)

    def GetItemParents(self, item_id):
        """Return a root-first list of parent item ids of `item_id`

        If `item_id` is an iterable, return a list of lists, corresponding to
        the sequence of items.
        """
        return self.one_or_many_to_many('GetItemParents', item_id)

    def GetItemSubs(self, item_id, depth=0):
        """itemId -> [(child_id,indent), ... ]"""
        return fold(self.intlist('GetItemSubs',depth,item_id))

    def GetItemText(self, item_id):
        """Text for `item_id` (or a list of strings if id is an iterable)"""
        return self.one_or_many('GetItemText', item_id, str)



    def GetItemType(self, item_id):
        """Type for `item_id` (or a list of types if id is an iterable)"""
        return self.one_or_many('GetItemType', item_id)

    def GetSelection(self):
        """Returns a list: [ type (1=items, 2=folders), selectedIds]"""
        res = [ map(int,line) for line in self('GetSelection') ]
        res[0] = res[0][0]
        return res

    def GetVersion(self):
        """Return the Ecco API protocol version triple (major, minor, rev#)"""
        return self.intlist('GetVersion')

    def NewFile(self):
        """Create a new 'Untitled' file, returning a session id"""
        time.sleep(0.01)    # give Ecco a chance to catch up
        return int(self('NewFile')[0][0])

    def OpenFile(self, pathname):
        """Open or switch to `pathname` and return a session ID

        If the named file was not actually opened (not found, corrupt, etc.),
        a ``ecco_dde.FileNotOpened`` error will be raised instead.
        """
        result = self(format([['OpenFile', pathname]]))[0]
        result = result and int(result[0]) or 0
        if not result: raise FileNotOpened(pathname)
        return result

    def PasteOLEItem(self, mode=OLEMode.Embed, item_id=None, data=()):#
        """Paste from  the clipboard, returning an item id
        
        If `item_id` is not None, the paste will go into that item.  `mode`
        should be ``OLEMode.Link`` or ``OLEMode.Embed`` (the default).  `data`,
        if supplied, should be a sequence of ``(folderid,value)`` pairs for the
        item to be initialized or updated with.
        """
        if item_id is None: item_id = ''
        return int(self('PasteOleItem', mode, item_id, *unfold(data))[0][0])

    # --- "Extended DDE Requests"

    def GetChanges(self, timestamp, folder_ids=()):
        """Get changes since `timestamp`, optionally restricted to `folder_ids`

        `timestamp` is an opaque value from Ecco itself, supplied as part of
        the return value from this method.  If `folder_ids` is supplied, only
        changes to those folders and the items in them are included. 
        Returns a triple of ``(nextstamp, items, folders)``, where
        ``nextstamp`` is the value that should be passed in to the next call to
        ``GetChanges()``, ``items`` is a list of items with changed text or
        folder values, and ``folders`` is a list of folders that have had items
        removed.  Note that due to the way Ecco processes change timestamps,
        not all listed items or folders may have actually changed since your
        last call to this method.
        """
        data = self('GetChanges', timestamp, *folder_ids)+[[],[]]
        return int(data[0][0]), map(int, data[1]), map(int, data[2])

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

    def GetPopupValues(self, folder_id):
        """Popup values for `folder_id` (or list of lists if id is iterable)"""
        return self.one_or_many_to_many('GetPopupValues', folder_id, str)

    def GetFolderOutline(self):
        """Return a list of ``(folderid, depth)`` pairs for the current file"""
        return fold(self.intlist("GetFolderOutline"))

    def GetViewColumns(self, view_id):
        """Folderids for `view_id` columns (`view_id` must be a single int)"""
        return self.intlist('GetViewColumns', view_id)

    def GetViewTLIs(self, view_id):
        """Return a list of ``(folder_id, itemlist)`` pairs for `view_id`"""
        rows = self('GetViewTLIs', view_id)
        for pos, row in enumerate(rows):
            rows[pos] = int(row.pop(0)), map(int, row)
        return rows

    def GetOpenFiles(self):
        """Return a list of session IDs for all currently-open files"""
        self.open() # ensure connect errors propagate
        try: return self.intlist("GetOpenFiles")
        except: return []

    def CreateView(self, name, folder_ids):
        assert folder_ids, "Must include at least one folder ID!"
        return self.intlist('CreateView', name, *folder_ids)[0]

    def GetFolderAutoAssignRules(self, folder_id):
        """Get list of strings defining auto-assign rules for `folder_id`"""
        self.open() # ensure connect errors propagate
        try: return self('GetFolderAutoAssignRules', folder_id)[0]
        except: return []

    def GetCurrentFile(self):
        """Return the session id of the active file"""
        self.open() # ensure connect errors propagate
        try: return int(self('GetCurrentFile')[0][0])
        except: return None

    def GetFileName(self, session_id):
        """Return the file name for the given session ID"""
        return self(format([['GetFileName', session_id]]))[0][0]

    # --- "DDE Pokes supported"

    def ChangeFile(self, session_id):
        """Switch to the designated `session_id`"""
        if self.GetCurrentFile()!=session_id:
            # Alas, this poke doesn't always work, at least not in my Ecco...
            self.poke('ChangeFile', session_id)
            # So we may have to use OpenFile instead:
            if self.GetCurrentFile()!=session_id:
                self.OpenFile(self.GetFileName(session_id))

    def CloseFile(self, session_id):
        """Close the designated session, *without* saving it"""
        self.assert_session(session_id)
        self.poke('CloseFile'); self.close() # force re-open for next access

    def CopyOLEItem(self, item_id):#
        """Copy the specified OLE item to the Windows clipboard"""
        self.poke('CopyOLEItem', item_id)

    def InsertItem(self, anchor_id, items, where=InsertLevel.Indent):
        """Insert item or items at `anchor_id` w/optional indent

        `where` should be ``InsertLevel.Indent``, ``InsertLevel.Outdent``,
        or ``InsertLevel.Same`` (default is ``Indent``).  `items` may be
        a single item id, or a sequence of items.  The items are moved to
        a point relative to `anchor_id`, which may be 0 to indicate that
        the item(s) are to become top-level.
        """
        if not hasattr(items, '__iter__'): items = [items]
        self.poke('InsertItem', anchor_id, where, *items)






    def RemoveItem(self, item_id):
        """Delete `item_id` (can be an iterable of ids)"""
        if hasattr(item_id, '__iter__'):
            self.poke('RemoveItem', *item_id)
        else:
            self.poke('RemoveItem', item_id)

    def SaveFile(self, session_id, pathname=None):
        """Save the designated session to `pathname`; fails if not current"""
        self.assert_session(session_id)
        if pathname:
            self.poke('SaveFile', pathname)
        else:
            self.poke('SaveFile')

    def SetFolderName(self, folder_id, name):
        """Set the name of `folder_id` to `name`"""
        self.poke('SetFolderName', folder_id, name)


    def SetFolderValues(self, item_ids, folder_ids, values):
        """Return folder values for specified folders and items

        `item_ids` can be a single item ID, or a sequence.  If it's a sequence,
        then `values` must be a sequence ordered by the input item ids.

        `folder_ids` can be a single folder ID, or a sequence.  If it's a
        sequence, the `values` for each item must be a sequence ordered by the
        input folder ids.

        If a single item ID and single folder ID are used, then `values` must
        be a single value.  If a single folder ID is used and multiple item
        IDs, then `values` must be a list of values, one value per item.  If
        multiple folder IDs are used, and only one item ID, then `values` must
        be a single list containing the values for that one item.

        In other words, depending on the target, you'll set either a value, a
        list of values (either different folders for one item, or one folder
        for different items), or a list of lists of values.
        """
        
        items, folders = cmd = [[], []]

        multi_folder = hasattr(folder_ids, '__iter__')
        multi_item = hasattr(item_ids, '__iter__')

        if multi_folder:
            folders.extend(folder_ids)
            if multi_item:
                if not hasattr(values, '__len__'):
                    values = list(values)
            else:
                values = [values]
        else:
            folders.append(folder_ids)
            if multi_item:
                values = [[v] for v in values]
            else:
                values = [[values]]

        if multi_item:
            items.extend(item_ids)
        else:
            items.append(item_ids)

        if len(values)!=len(items):
            raise ValueError("Length mismatch between item_ids and values")

        fc = len(folders)
        for v in values:
            if not hasattr(v, '__len__'):
                v = list(v)
            if len(v)!=fc:
                raise ValueError("Length mismatch between folder_ids and values")
            cmd.append(v)

        if self.connection is None:
            self.open()
        self.connection.Poke('SetFolderValues', format(cmd))



    def SetItemText(self, item_id, text=None):
        """Set the text of `item_id` to `text` (or a dictionary of item->text)

        If `text` is None, `item_id` must be a dictionary mapping item ID's to
        text values.  Otherwise, `item_id` should be a single item ID, and
        `text` is the text to set.
        """
        if text is None:
            self.poke('SetItemText', *unfold(item_id.items()))            
        else:
            self.poke('SetItemText', item_id, text)
        
    def ShowPhoneBookItem(self, item_id, only=True):
        """Show the specified item in the phonebook view

        Item or one of its parents must be in the Phonebook folder.  If `only`
        is true, the phonebook will show only this item.  Otherwise, the item
        is added to the current list of phonebook items.

        This also switches to the phonebook view, if it's not already visible.
        """
        self.poke('ShowPhoneBookItem', item_id, int(bool(only)))

    # --- "Extended DDE Pokes"

    def ChangeView(self, view_id):
        """Display the specified view"""
        self.poke('ChangeView', view_id)

    def AddCompView(self, view_id):
        """Add `view_id` as a composite view to the current view"""
        self.poke('AddCompView', view_id)

    def RemoveCompView(self, view_id):
        """Remove the specified view from the current view's composite views"""
        self.poke('RemoveCompView', view_id)





    def SetCalDate(self, date):
        """Display `date` in the calendar (only if calendar is already visible)

        `date` may be a ``datetime.date`` or ``datetime.datetime`` instance, or
        any other object with a ``strftime()`` method.  Otherwise, it should be
        a string already formatted to Ecco's date format.
        """
        self.poke('SetCalDate', format_date(date))

    def DeleteView(self, view_id):
        """Delete the specified view (`view_id` must be a single int)"""
        self.poke('DeleteView', view_id)

    #AddFileToMenu FilePath IconID

    def AddColumnToView(self, view_id, folder_id):
        """Add the specified folder(s) as column(s) of `view_id`"""
        self.poke_one_or_many('AddColumnToView', folder_id, view_id)

    def AddFolderToView(self, view_id, folder_id):
        """Add the specified folder(s) to contents of `view_id`"""
        self.poke_one_or_many('AddFolderToView', folder_id, view_id)

    def poke_one_or_many(self, cmd, ob, *args):
        if hasattr(ob, '__iter__') and not isinstance(ob, basestring):
            self.poke(cmd, *args+tuple(ob))
        else:
            self.poke(cmd, *args+(ob,))













