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

.. EccoDDE developer's guide: http://peak.telecommunity.com/DevCenter/EccoDDE#toc
.. _PEAK mailing list: http://www.eby-sarna.com/mailman/listinfo/peak/

.. _toc:
.. contents: **Table of Contents**


-----------------
Developer's Guide
-----------------

To talk to Ecco, you will use an ``EccoDDE`` instance::

    >>> from ecco_dde import EccoDDE
    >>> api = EccoDDE()

The ``EccoDDE`` constructor accepts the following keyword-only arguments, which
are used only if an initial attempt to contact Ecco fails::

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
    ['ItemText', 'OLELink', ...]

An item is of type ``ItemText`` if it's a normal text item, and ``OLELink`` if
it's an OLE item (e.g., one created using ``PasteOLEItem()``.


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

You can set folder values using ``SetFolderValues()``, with either individual
item ids and folder ids, or sequences thereof::

    >>> api.SetFolderValues(an_item, f1, 1)
    >>> api.SetFolderValues(another_item, [f1,f3], [1,'some text'])
    >>> api.SetFolderValues([an_item,another_item], f4, ['20010101','20020202'])

    XXX need is+fs example

And you can retrieve the values in a similar fashion, using
``GetFolderValues()``::

    >>> api.GetFolderValues(another_item, f1)
    '1'

    >>> api.GetFolderValues(another_item, [f3, f1])
    ['some text', '1']

    >>> api.GetFolderValues([an_item,another_item], f4)
    ['20010101', '20020202']

    XXX need is+fs example

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

GetItemParents - one, many
GetItemSubs - depth 0, depth ?

InsertItem - one, many

    >>> from ecco_dde import InsertLevel
    >>> dir(InsertLevel)
    ['Indent', 'Outdent', 'Same', ...]

RemoveItem - one, many


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

And you can query a view's folders using ``GetViewFolders()``, either with a
single view ID::

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

    >>> api.AddFolderToView(view, [f1, f3])

    >>> api.GetViewFolders(view) == date_folders + [popup, f1, f3]
    True


AddColumnToView - one, many     XXX - vis
GetViewColumns - one, many

GetViewTLIs

ChangeView          XXX - vis
AddCompView         XXX - vis
RemoveCompView      XXX - vis

DeleteView - one, many


Miscellaneous APIs
==================

The ``GetVersion()`` method returns the Ecco API version number::

    >>> api.GetVersion()
    [2, 8, 0]


GetChanges          XXX

GetSelection        XXX


SetCalDate          XXX - vis



Date and time formatting::

    >>> from ecco_dde import format_date, format_datetime

    >>> from datetime import datetime
    >>> dt = datetime(2008, 3, 31, 17, 53, 46)

    >>> format_date(dt)
    '20080331'

    >>> format_datetime(dt)
    '200803311753'

    >>> format_date(27)     # objects w/out strftime pass thru
    27

    >>> format_datetime(99)
    99

ShowPhoneBookItem
CopyOLEItem         XXX
PasteOLEItem        XXX

    >>> from ecco_dde import OLEMode
    >>> dir(OLEMode)
    ['Embed', 'Link', ...]


Wrap-up::

    >>> api.close_all()
    >>> api.close()

    >>> from shutil import rmtree
    >>> rmtree(tmpdir)

