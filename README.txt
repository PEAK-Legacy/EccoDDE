============================
Accessing Ecco using EccoDDE
============================

EccoDDE is a very thin wrapper over the DDE API of the Ecco Pro personal
information manager for Windows.  Unlike other Python interfaces to Ecco,
it does not provide any higher-level objects to represent items, folders, etc.,
but instead permits you to create whatever higher-level objects suit your
particular application.

Also, EccoDDE uses the original Ecco API names for its methods, so that you can
use the Ecco API reference as a rough guide to EccoDDE.  Some methods have
enhanced functionality that you can access by using different argument types,
but even these are nearly always just exposing capabilities of the underlying
Ecco API, rather than doing any Python-specific wrapping.

The main value-add that EccoDDE provides over ad-hoc scripting is robustness.
EccoDDE can transparently launch Ecco if it's not started, and it avoids many
subtle quoting and line-termination problems that you'd run into when writing
an interface from scratch.  EccoDDE also has an automated test suite, so that
any future additions to the library won't break current functionality.

This library requires the PyWin32 package, but does not automatically install
it, due to compatibility issues.  You must manually install PyWin32 before
using EccoDDE.

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

The main objects that EccoDDE provides are the ``EccoDDE`` class, and the
``FolderType`` and ``ItemType`` enumerations::

    >>> from ecco_dde import EccoDDE, FolderType, ItemType

    >>> dir(FolderType)
    ['CheckMark', 'Date', 'Number', 'PopUpList', 'Text', ...]

    >>> dir(ItemType)
    ['ItemText', 'OLELink', ...]

To talk to Ecco, you will use an ``EccoDDE`` instance::

    >>> api = EccoDDE()

The ``EccoDDE`` constructor accepts the following keyword arguments, which
are only used if an initial attempt to contact Ecco fails::

filename
    The filename to launch to run Ecco.  If ``None`` or not supplied, the
    Windows registry will be consulted to find the shell-open command for
    Ecco files.

retries
    The number of times to try connecting to Ecco after attempting to launch
    it.  (10 by default)

sleep
    The number of seconds to sleep between connection attempts (1 by default)

Note, too, that creating an ``EccoDDE`` instance does not immediately launch or
connect to Ecco.  You can explicitly call the ``open()`` method, if you like,
but it will also be called automatically whenever necessary.  The ``close()``
method can be used to shut down the connection when it's not in use.

Some usage examples::

    >>> api.GetVersion()
    [2, 8, 0]

    >>> api.close_all()     # close everything else first
    
    >>> session = api.NewFile()

    >>> api.GetOpenFiles() == [session]
    True

    >>> api.GetCurrentFile() == session
    True

    >>> api.GetFileName(session)
    '<Untitled>'
    
    >>> api.GetViewNames(api.GetViews())
    ['Calendar', 'PhoneBook']

    >>> api.GetViewNames()
    [('Calendar', 2), ('PhoneBook', 3)]

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

    >>> api.GetFoldersByName('Appointments')
    [4]

    >>> date_folders = api.GetFoldersByType(FolderType.Date)

    >>> for folder in date_folders:
    ...     print api.GetFolderName(folder)
    Phone / Time Log
    Appointments
    Done
    Start Dates
    Due Dates
    To-Do's
    Recurring Note Dates

    >>> view = api.CreateView('All Dates', date_folders)

    >>> api.GetViewFolders(view) == date_folders
    True


...and 22 more methods to test here, plus 13 more not implemented yet...


    >>> api.CloseFile(session)

    >>> api.close()



-------------------
Internals and Tests
-------------------

XXX
