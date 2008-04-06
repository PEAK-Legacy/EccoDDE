============================
Scripting Ecco using EccoDDE
============================

EccoDDE is a very thin wrapper over the DDE API of the Ecco Pro personal
information manager for Windows.  Unlike other Python interfaces to Ecco,
it does not provide any higher-level objects to represent items, folders, etc.,
but instead permits you to create whatever higher-level objects suit your own
particular application.

Also, EccoDDE uses the original Ecco API names for its methods, so that you can
use the Ecco API reference as a rough guide to EccoDDE.  Some methods have
enhanced functionality that you can access by using different argument types,
but even these are nearly always just exposing capabilities of the underlying
Ecco API, rather than doing any Python-specific wrapping.  48 of Ecco's 49 API
calls are implemented.  (The 49th, ``AddFileToMenu``, does not have appear to
be documented anywhere on the 'net.)

The main value-add that EccoDDE provides over writing your own ad-hoc interface
is robustness.  EccoDDE can transparently launch Ecco if it's not started, and
it avoids many subtle quoting and line-termination problems that you'd run into
when writing an interface from scratch.  EccoDDE also has an automated test
suite, so that any future additions to the library won't break current
functionality.

This library requires the PyWin32 package, but does not automatically install
it, due to it not being compatible with easy_install at this time.  You must
manually install PyWin32 before using EccoDDE.

For complete EccoDDE documentation, please consult the `EccoDDE developer's
guide`_.  Questions, comments, and bug reports for this package should be
directed to the `PEAK mailing list`_.

.. _Trellis: http://pypi.python.org/pypi/Trellis
.. _Trellis tutorial: http://peak.telecommunity.com/DevCenter/Trellis

.. _EccoDDE developer's guide: http://peak.telecommunity.com/DevCenter/EccoDDE#toc
.. _PEAK mailing list: http://www.eby-sarna.com/mailman/listinfo/peak/

.. _toc:
.. contents:: **Table of Contents**


-----------------
Developer's Guide
-----------------

To talk to Ecco, you will use an ``EccoDDE`` instance::

    >>> from ecco_dde import EccoDDE
    >>> api = EccoDDE()

The ``EccoDDE`` constructor accepts the following keyword-only arguments, which
are used only if an initial attempt to contact Ecco fails:

filename
    The filename to launch (with ``os.startfile()``) to run Ecco.  If ``None``
    or not supplied, the Windows registry will be consulted to find the
    shell-open command registered for Ecco files.

retries
    The number of times to try connecting to Ecco after attempting to launch
    it.  (10 by default)

sleep
    The number of seconds to sleep between connection attempts (1 by default)

Note, too, that creating an ``EccoDDE`` instance does not immediately launch or
connect to Ecco.  You can explicitly call the ``open()`` method, if you like,
but it will also be called automatically whenever necessary.  The ``close()``
method can be used to shut down the connection when it's not in use.  If you
use an instance after closing it, it will automatically be re-opened, which
means you can (and probably should) close the connection when you won't be
using it for a while.


Working With Files and Sessions
===============================

The ``close_all()`` method closes all currently-open files::

    >>> api.close_all()     # close any files currently open in Ecco

    >>> api.GetOpenFiles()
    []

``NewFile()`` creats a new, untitled file, returning a session ID::

    >>> session = api.NewFile()

Which then will be visible in ``GetOpenFiles()`` (a list of the active session
IDs), and as the ``GetCurrentFile()`` (which returns the active session ID)::

    >>> api.GetOpenFiles() == [session]
    True

    >>> api.GetCurrentFile() == session
    True

The newly created file will be named '<Untitled>'::

    >>> api.GetFileName(session)
    '<Untitled>'

Until it is saved::

    >>> from tempfile import mkdtemp
    >>> tmpdir = mkdtemp()

    >>> import os
    >>> testfile = os.path.join(tmpdir, 'testfile.eco')

    >>> os.path.exists(testfile)
    False

    >>> api.SaveFile(session, testfile)

    >>> os.path.exists(testfile)
    True

    >>> api.GetFileName(session)
    '...\\testfile.eco'

Once a session has a filename, it can be saved without specifying the name::

    >>> api.SaveFile(session)

And the ``CloseFile()`` and ``OpenFile()`` APIs work much as you would expect::

    >>> api.CloseFile(session)

    >>> session = api.OpenFile(testfile)

    >>> api.GetOpenFiles() == [session]
    True

    >>> api.GetCurrentFile() == session
    True

And you can also use the ``ChangeFile()`` API to switch to a given session::

    >>> session2 = api.NewFile()
    >>> session2 == api.GetCurrentFile()
    True

    >>> api.SaveFile(session2, os.path.join(tmpdir, 'test2.eco'))

    >>> api.ChangeFile(session)
    >>> session == api.GetCurrentFile()
    True
    >>> session2 == api.GetCurrentFile()
    False

    >>> api.ChangeFile(session2)
    >>> session == api.GetCurrentFile()
    False
    >>> session2 == api.GetCurrentFile()
    True

Note, by the way, that you can only close or save a file if it is the current
session::

    >>> api.SaveFile(session)
    Traceback (most recent call last):
      ...
    StateError: Attempt to close or save inactive session

    >>> api.CloseFile(session)
    Traceback (most recent call last):
      ...
    StateError: Attempt to close or save inactive session

    >>> api.CloseFile(session2)


Working With Folders
====================


Listing and Looking Up Folders
------------------------------

The ``GetFolderOutline()`` method returns a list of ``(depth, id)`` tuples
describing the folder outline of the current Ecco file, while the
``GetFolderName()`` and ``GetFolderType()`` methods return the name or type
for a given folder ID::

    >>> for folder, depth in api.GetFolderOutline():
    ...     print "%-30s %02d" % (
    ...         '   '*depth+api.GetFolderName(folder),api.GetFolderType(folder)
    ...     )
    Ecco Folders                   01
       PhoneBook                   01
          Mr./Ms.                  04
          Job Title                04
          Company                  04
          Address 1 - Business     04
          Address 2 - Business     04
          City - Business          04
          State - Business         04
          Zip - Business           04
          Country - Business       04
          Work #                   04
          Home #                   04
          Fax #                    04
          Cell #                   04
          Alt #                    04
          Address 1 - Home         04
          Address 2 - Home         04
          City - Home              04
          State - Home             04
          Zip - Home               04
          Country - Home           04
          Phone / Time Log         02
          E-Mail                   04
       Appointments                02
       Done                        02
       Start Dates                 02
       Due Dates                   02
       To-Do's                     02
       Search Results              01
       New Columns                 01
          Net Location             04
          Recurring Note Dates     02

The values returned by ``GetFolderType()`` are available as constants in the
``FolderType`` enumeration class::

    >>> from ecco_dde import FolderType

    >>> dir(FolderType)
    ['CheckMark', 'Date', 'Number', 'PopUpList', 'Text', ...]

    >>> FolderType.CheckMark
    1

Which makes it convenient to fetch a list of folder ids based on folder type,
using the ``GetFoldersByType()`` method::

    >>> date_folders = api.GetFoldersByType(FolderType.Date)

Both ``GetFolderName()`` and ``GetFolderType()`` will return multiple values if
their input is an iterable::

    >>> for name in api.GetFolderName(date_folders):    # accepts multiples
    ...     print name
    Phone / Time Log
    Appointments
    Done
    Start Dates
    Due Dates
    To-Do's
    Recurring Note Dates

    >>> api.GetFolderType(date_folders)
    [2, 2, 2, 2, 2, 2, 2]

You can also find the folders by name, using ``GetFoldersByName()``::

    >>> api.GetFoldersByName('Appointments')
    [4]

(Note that this method always returns a list of ids, since more than one folder
can have the same name.)


Creating And Managing Folders
-----------------------------

The ``CreateFolder()`` API can be used to create a single folder::

    >>> f1 = api.CreateFolder('Test Folder 1')

By default, it's created as a checkmark folder:

    >>> api.GetFolderType(f1) == FolderType.CheckMark
    True

But you can also specify a type explicitly::

    >>> popup = api.CreateFolder('A popup folder', FolderType.PopUpList)
    >>> api.GetFolderType(popup) == FolderType.PopUpList
    True

``CreateFolder()`` can also create multiple folders at once, using a dictionary
mapping names to folder types::

    >>> d = api.CreateFolder(
    ...     {'folder 3':FolderType.Text, 'folder 4':FolderType.Date}
    ... )

And the return value is a dictionary mapping the created folder names to their
folder ids::

    >>> d
    {'folder 4': ..., 'folder 3': ...}

    >>> f3 = d['folder 3']
    >>> f4 = d['folder 4']

    >>> api.GetFolderName(f3)
    'folder 3'

    >>> api.GetFolderType(f4)==FolderType.Date
    True

You can also rename an existing folder using ``SetFolderName()``::

    >>> api.SetFolderName(f4, 'A Date Folder')
    >>> api.GetFolderName(f4)
    'A Date Folder'


Pop-ups and Auto-assign Rules
-----------------------------

You can retrieve a pop-up folder's values using ``GetPopupValues()``::

    >>> api.GetPopupValues(popup)
    []

If you pass in an iterable of folder ids, you will get back a list of lists of
pop-up values::

    >>> api.GetPopupValues([popup])
    [[]]

At the moment, our example popup folder doesn't have any values.  Creating or
modifying an item with values for this folder will add some::

    >>> i1 = api.CreateItem('Test 1', [(popup, 'Blue')])
    >>> i2 = api.CreateItem('Test 2', [(popup, 'Red')])

    >>> api.GetPopupValues(popup)
    ['Blue', 'Red']

And the added values will stay around even if we delete the items::

    >>> api.RemoveItem(i1)
    >>> api.RemoveItem(i2)

    >>> api.GetPopupValues([popup])
    [['Blue', 'Red']]

You can query a folder's auto-assign rules (if any) using
``GetFolderAutoAssignRules()`` (which will only accept one folder id, btw)::

    >>> api.GetFolderAutoAssignRules(api.GetFoldersByName('Net Location')[0])
    ['http:#']

By the way, there is no way to programmatically delete an existing folder,
change its type, or add/change its auto-assignment rules.  These actions can
only be done through the Ecco UI.


Working With Items
==================


Creating and Inspecting Items
-----------------------------

As we saw above, you can create items from text, and an optional sequence of
``(folderid,value)`` pairs, passed to ``CreateItem()``::

    >>> an_item = api.CreateItem('An item')

    >>> another_item = api.CreateItem('Another item', [(popup, 'Red')])

You can find out an item's text, folders, or type using ``GetItemText()``,
``GetItemFolders()``, and ``GetItemType()`` respectively::

    >>> api.GetItemText(an_item)
    'An item'

    >>> api.GetItemFolders(another_item)==[popup]
    True

    >>> api.GetItemType(an_item)
    1

Or you can pass in multiple item ids as a sequence, to get back a list of
strings, item types, or lists of folder ids::

    >>> both = [an_item, another_item]

    >>> api.GetItemText(both)
    ['An item', 'Another item']

    >>> api.GetItemFolders(both)==[[], [popup]]
    True

    >>> api.GetItemType(both)
    [1, 1]

By the way, there are item type constants declared in the ``ItemType`` class::

    >>> from ecco_dde import ItemType
    >>> dir(ItemType)
    ['ItemText', 'OLE', ...]

An item is of type ``ItemText`` if it's a normal text item, and ``OLE`` if
it's an OLE item (e.g., one created using ``PasteOLEItem()``).


Querying and Updating Items/Values
----------------------------------

The ``GetFolderItems()`` method returns a list of item ids for a given folder
id::

    >>> api.GetFolderItems(popup) == [another_item]
    True

It can also accept optional extra arguments to specify sort and search options,
as described in the Ecco API documents::

    >>> api.GetFolderItems(popup, 'EQ', 'Red') == [another_item]
    True

    >>> api.GetFolderItems(popup, 'EQ', 'Blue')
    []

You can set and retrieve folder values using ``SetFolderValues()`` and
``GetFolderValues``, using either individual item ids and folder ids, or
sequences thereof::

    >>> api.SetFolderValues(an_item, f1, 1)
    >>> api.GetFolderValues(an_item, f1)
    '1'

    >>> api.SetFolderValues(another_item, [f1,f3], [1,'some text'])
    >>> api.GetFolderValues(another_item, [f3, f1])
    ['some text', '1']

    >>> api.SetFolderValues(
    ...     [an_item, another_item], f4, ['20010101', '20020202']
    ... )
    >>> api.GetFolderValues([an_item,another_item], f4)
    ['20010101', '20020202']

    >>> api.SetFolderValues(
    ...     [an_item, another_item], [popup, f3],
    ...     [['Blue', 'text a'], ['Red', 'text b']]
    ... )
    >>> api.GetFolderValues([an_item, another_item], [popup, f3])
    [['Blue', 'text a'], ['Red', 'text b']]

Now let's retrieve some items sorted by their item text (ascending and
descending)::

    >>> api.GetFolderItems(f1, 'ia') == [an_item, another_item]
    True

    >>> api.GetFolderItems(f1, 'id') == [another_item, an_item]
    True

You can also change items' text using ``SetItemText()``, either by passing in a
single item id and text::

    >>> api.SetItemText(an_item, 'A')
    >>> api.GetItemText(an_item)
    'A'

Or by passing a dictionary mapping item id's to the desired text::

    >>> api.SetItemText({an_item:'1', another_item:'2'})
    >>> api.GetItemText([an_item, another_item])
    ['1', '2']


Item Hierarchy, Relocation, and Removal
---------------------------------------

The ``GetItemParents()`` method returns a (possibly-empty) list of parent item
ids for one or more items::

    >>> api.GetItemParents(an_item)
    []

    >>> api.GetItemParents([an_item, another_item])
    [[], []]

Hm, no parents so far, so let's rearrange the hierarchy using ``InsertItem()``.
For that, we need to know the "anchor item" (a parent or sibling), and the item
or items to be placed relative to it.  We also need to specify an insertion
level::

    >>> from ecco_dde import InsertLevel
    >>> dir(InsertLevel)
    ['Indent', 'Outdent', 'Same', ...]

The default level of ``Indent``places the targeted items as the first
child(ren) of the anchor item, while ``Outdent`` places them as siblings of the
anchor item's parent.  ``Same`` makes the items a sibling of the existing item.
Let's begin by creating a new item, then make our existing items children of
it::

    >>> i3 = api.CreateItem('Item 3')
    >>> api.InsertItem(i3, [an_item, another_item])

    >>> api.GetItemParents([an_item, another_item]) == [[i3], [i3]]
    True

The ``GetItemSubs()`` method can be used to retrieve a sequence of
``(depth,id)`` pairs, describing the children of an item::

    >>> api.GetItemSubs(i3) == [(1, another_item), (1, an_item)]
    True

Notice that the items here are listed in reverse order from the order we
added them.  This is because the default ``InsertLevel.Indent`` mode inserts
children at the head of the parent's list of children.

By default, ``GetItemSubs()`` returns a list of all children, to any depth::

    >>> api.InsertItem(an_item, another_item)

    >>> api.GetItemParents(another_item) == [i3, an_item]
    True

    >>> api.GetItemSubs(i3) == [(1, an_item), (2, another_item)]
    True

But it can be given a depth argument in order to prune the returned tree::

    >>> api.GetItemSubs(i3, 1) == [(1, an_item)]
    True

Notice, by the way, that ``GetItemParents()`` returns parents in high-to-low
order, i.e., first the top-level item id, and the item's immediate parent last.

Also notice that ``InsertItem()`` might better be called "move item", since it
relocates the item to the specified location.  You can detach an item from
any existing parent using zero for the anchor id, and of course its children
will move with it::

    >>> api.InsertItem(0, an_item)

    >>> api.GetItemParents(another_item) == [an_item]
    True

    >>> api.GetItemSubs(i3)
    []

Using a depth of ``InsertLevel.Same``, we can now put ``i3`` after
``another_item``::

    >>> api.InsertItem(another_item, i3, InsertLevel.Same)

    >>> api.GetItemSubs(an_item) == [(1, another_item), (1, i3)]
    True

And let's add a couple more items so we can see how ``InsertLevel.Outdent``
works::

    >>> i4 = api.CreateItem('Item 4')
    >>> i5 = api.CreateItem('Item 5')

    >>> api.InsertItem(another_item, i4)
    >>> api.InsertItem(i4, i5, InsertLevel.Outdent)
    >>> api.GetItemSubs(an_item) == [(1,another_item), (2,i4), (1,i5), (1,i3)]
    True

As you can see, outdenting places an item in the next spot after the anchor
item's parent.

We can now delete some of our unneeded items with ``RemoveItem``, which works
with either a single item id, or an iterable of them::

    >>> api.RemoveItem(i4)
    >>> api.GetItemSubs(an_item) == [(1, another_item), (1, i5), (1, i3)]
    True

    >>> api.RemoveItem([i3,i5])
    >>> api.GetItemSubs(an_item) == [(1, another_item)]
    True

(Note, by the way, that ``RemoveItem`` will also delete an item's children, if
any exist.)


Working With Views
==================

The ``GetViews()`` method returns a list of view IDs for the current file::

    >>> api.GetViews()
    [2, 3]

And the ``GetViewNames()`` method returns a list of names, given a list of
view IDs::

    >>> api.GetViewNames([2, 3])
    ['Calendar', 'PhoneBook']

Or one name, if given a single view ID::

    >>> api.GetViewNames(2)
    'Calendar'

Or a sequence of ``(name, view_id)`` pairs for all views in the current file,
if not given an argument::

    >>> api.GetViewNames()
    [('Calendar', 2), ('PhoneBook', 3)]

You can create a new view by passing a name and a non-empty list of folders to
the ``CreateView()`` method::

    >>> view = api.CreateView('All Dates', date_folders)

    >>> api.GetViewNames()
    [('Calendar', 2), ('PhoneBook', 3), ('All Dates', 5)]


Folders and TLIs
----------------

You can query a view's folders using ``GetViewFolders()``, either with a single
view ID::

    >>> api.GetViewFolders(view) == date_folders
    True

or with a list of view IDs, to get a list of lists of folders for each view::

    >>> api.GetViewFolders([view]) == [date_folders]
    True

You can also add folders to a view (although you can't remove them)::

    >>> api.AddFolderToView(view, popup)

    >>> api.GetViewFolders(view) == date_folders + [popup]
    True

You can add multiple folders by passing a sequence as the second argument,
instead of a single folder id::

    >>> view2 = api.CreateView('Another View', [f1])
    
    >>> api.AddFolderToView(view2, [f1, f3])

    >>> api.GetViewFolders(view2) == [f1, f3]
    True

And the ``GetViewTLIs`` method returns a list of ``(folderid, itemids)`` pairs,
giving you each folder in the view and its corresponding top-level items::

    >>> api.GetViewTLIs(view2) == [(f1, [an_item]), (f3, [an_item])]
    True

By the way, you can pass a sequence of views to ``GetViewFolders()``::

    >>> api.GetViewFolders([view, view2]) == [
    ...     date_folders+[popup], [f1,f3]
    ... ]
    True


Columns
-------

By default, any non-checkmark folders added to a view programmatically will
also be added to the view as columns, with the most-recently added folders
first::

    >>> api.GetViewColumns(view2) == [f3]
    True

But you can also add columns explicitly, using ``AddColumnToView`` to specify
one or many folders::

    >>> api.AddColumnToView(view2, f1)
    >>> api.GetViewColumns(view2) == [f1, f3]
    True

    >>> api.AddColumnToView(view2, [f4, popup])
    >>> api.GetViewColumns(view2) == [f4, popup, f1, f3]
    True
    
And please note, by the way, that ``GetViewColumns()`` does NOT work with
multiple view ids.  In testing, Ecco honors the format of the API by returning
multiple lists, but the lists are all empty except for the first one.  So if
you need the column information for multiple views, you'll need to do the
looping yourself.


Selecting, Composing, and Deleting Views
----------------------------------------

The ``ChangeView()`` method lets you select which view is displayed in Ecco::

    >>> api.ChangeView(view)
    >>> api.ChangeView(view2)

Unfortunately, there's no way to get the current view ID, and thus no way of
knowing whether this actually worked.  Similarly, there's no way to find out
what split-screen views are active, so even though you can add and remove up to
3 additional ("composite") views on the current view, there's no way to see if
it's working, either::

    >>> api.AddCompView(2)      # add the calendar
    >>> api.AddCompView(3)      # and the phonebook
    >>> api.AddCompView(view)   # and the other view...
    
    >>> api.RemoveCompView(2)       # take off the calendar
    >>> api.RemoveCompView(3)       # and the phonebook
    >>> api.RemoveCompView(view)    # and the other view...

The ``DeleteView()`` API removes a view from the current file::

    >>> api.GetViewNames()
    [('Calendar', 2), ('PhoneBook', 3), ('All Dates', 5), ('Another View', 6)]

    >>> api.DeleteView(view)
    >>> api.GetViewNames()
    [('Calendar', 2), ('PhoneBook', 3), ('Another View', 6)]

(Note that although the Ecco API docs claim that this function can be given
multiple view IDs, in practice it crashes Ecco.)


Miscellaneous APIs
==================

``GetVersion()``
----------------

The ``GetVersion()`` method returns the Ecco API version number::

    >>> api.GetVersion()
    [2, 8, 0]


``GetChanges()``
----------------

The ``GetChanges()`` API returns a 3-tuple of ``(nextstamp, items, folders)``,
where `nextstamp` is a timestamp to be passed back in the next time you call
``GetChanges()``, `items` are the item ids that were changed or created since
the last call, and `folders` are the folder ids of folders that had items
removed from them since the last call.

Initially, you should pass in a timestamp of 0, to get the ball rolling::

    >>> nextstamp, items, folders = api.GetChanges(0)

This will basically return all changes since the file was created, and a
timestamp you can use to call again.  If no changes have occurred since that
stamp, you'll get empty lists and the same timestamp again, the next time you
call::

    >>> api.GetChanges(nextstamp) == (nextstamp, [], [])
    True

Before using this, please read the Ecco DDE API docs carefully regarding how
the timestamp resolution works and when/why you'll get duplicate change
notices, as well as when/why certain notices will be delayed.  This is not
necessarily a suitable method for staying in sync "live" with an active Ecco
file, as you can miss certain kinds of changes for up to an hour before they'll
appear in this method's results.

``GetChanges()`` can optionally take a second argument, that lists the folders
to be watched.  Any items that aren't in those folders won't be included in
the results, but the folders themselves will be listed in the `folders` return
if items were removed from them during the relevant timeframe.


OLE Copy/Paste
--------------

The ``CopyOLEItem(item_id)`` method copies an OLE item to the clipboard.  To
work, the item must have a ``GetItemType()`` of ``ItemType.OLE``.

``PasteOLEItem(mode, optional_id, optional_data)`` pastes an OLE object from
the Windows clipboard, either into an existing item, or creating a new item
and returning its id (if `optional_id` is omitted or ``None``).  The first
argument is an ``OLEMode``, either ``Embed`` or ``Link``::

    >>> from ecco_dde import OLEMode
    >>> dir(OLEMode)
    ['Embed', 'Link', ...]

And the third argument, if supplied, should be a sequence of ``(folderid,val)``
pairs, that will be used to initialize or update the item.

(Unfortunately, there isn't any straightforward way to demo/test these API
calls in this document, as you must have an OLE object either in the clipboard
or in an existing Ecco file.  If you experience any problems using them, please
let me know how to reproduce the problem.  Thanks.)


``ShowPhoneBookItem()``
-----------------------

The ``ShowPhoneBookItem()`` method switches to the phonebook view and displays
the specified item.  If you pass a false value as the second argument, the
specified item will be added to the current search results in the phonebook
view.  Otherwise, the search results are cleared first::

    >>> pb, = api.GetFoldersByName('PhoneBook')
    >>> api.SetFolderValues([an_item, another_item], pb, [1,1])
    
    >>> api.InsertItem(0, another_item) # make this a top-level item

    >>> api.ShowPhoneBookItem(an_item)              # display an item...
    >>> api.ShowPhoneBookItem(another_item, False)  # then add another

Note, by the way, that the specified item must either be in the ``Phonebook``
folder, or have a parent that is.  Otherwise, an error will be raised.


``GetSelection()``
------------------

The ``GetSelection()`` method returns a ``[kind, ids]`` pair, where `ids` is a
list of folder or item ids, and `kind` is a ``SelectionType`` constant
indicating whether the ids are items or folders::

    >>> from ecco_dde import SelectionType
    >>> dir(SelectionType)
    ['Folders', 'Items', 'Nothing', ...]

In the previous section, we navigated to ``another_item``, so the current
selection should reflect that::

    >>> api.GetSelection() == [SelectionType.Items, [another_item]]
    True


``SetCalDate()``
----------------

The ``SetCalDate()`` method lets you set the calendar to the given date,
provided that you change to the calendar view, and use an appropriately
formatted date::

    >>> api.ChangeView(2)   # show the calendar
    >>> api.SetCalDate('20080301')


Date Conversion
===============

If you need to convert a Python date or datetime value to an Ecco DDE string,
use these functions::

    >>> from ecco_dde import format_date, format_datetime

    >>> from datetime import datetime
    >>> dt = datetime(2008, 3, 31, 17, 53, 46)

    >>> format_date(dt)
    '20080331'

    >>> format_datetime(dt)
    '200803311753'

These functions are not particularly bright, however, and will pass through
anything you give them that doesn't have a ``strftime()`` method::

    >>> format_date(27)     # objects w/out strftime pass thru
    27
    >>> format_datetime(99)
    99


Conclusion
==========

This is just a cleanup section where we close all open Ecco files, close our
DDE connection, and delete the temporary directory we used for saving test
files.  You probably don't want to do these things in your application, except
maybe closing the DDE connection::

    >>> api.close_all()     # close all files open in Ecco
    >>> api.close()         # close the DDE connection

    >>> from shutil import rmtree   # and wipe out the temp dir.
    >>> rmtree(tmpdir)

