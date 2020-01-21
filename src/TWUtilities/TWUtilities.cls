VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWUtilities"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class contains global methods and properties, and public enums and types.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                            As String = "TWUtilities"

Private Const ConfigSettingFontBold                 As String = "&FontBold"
Private Const ConfigSettingFontItalic               As String = "&FontItalic"
Private Const ConfigSettingFontName                 As String = "&FontName"
Private Const ConfigSettingFontSize                 As String = "&FontSize"
Private Const ConfigSettingFontStrikethrough        As String = "&FontStrikethrough"
Private Const ConfigSettingFontUnderline            As String = "&FontUnderline"

' command line switch specifying the log filename
Private Const SwitchLogFilename                     As String = "log"

' command line switch specifying the loglevel
Private Const SwitchLogLevel                        As String = "loglevel"

'@================================================================================
' Enums
'@================================================================================

''
' Used in the <code>ITheme</code> interface to indicate how themed controls should appear.
'
' @param AppearanceFlat
'   Themed controls should have a flat appearance.
' @param Appearance3D
'   Themed controls should have a 3D appearance.
'
'@/
Public Enum AppearanceSettings
    AppearanceFlat = 0
    Appearance3d = 1
End Enum

''
' Used in the <code>ITheme</code> interface to indicate the border style for themed controls.
'
' @param BorderStyleNone
'   Themed controls should have no border.
' @param BorderStyleSingle
'   Themed controls should have a single pixel border.
'
'@/
Enum BorderStyleSettings
    BorderStyleNone = 0
    BorderStyleSingle = 1
End Enum

''
' Used in the <code>CollectionChangeEventData</code> to indicate the nature of the change.
'
' @param CollItemAdded
'   An item has been added to the collection.
' @param CollItemRemoved
'   An item has been removed from the collection.
' @param CollItemChanged
'   An item in the collection has changed.
' @param CollOrderChanged
'   The order of items in the collections has changed.
' @param CollCollectionCleared
'   All items have been removed from the collection.
'
'@/
Public Enum CollectionChangeTypes
    CollItemAdded = 1
    CollItemRemoved
    CollItemChanged
    CollOrderChanged
    CollCollectionCleared
End Enum

''
' Used in the <code>ChangeEventData</code> for <code>ConfigStoreProvider</code> objects
'   to indicate the nature of the change.
'
' @param ConfigDirty
'   A change has been made to the configuration data.
' @param ConfigClean
'   The <code>ConfigStoreProvider</code> object has been saved.
'
'@/
Public Enum ConfigChangeTypes
    ConfigDirty = 1
    ConfigClean
End Enum

''
' Specifies how the location of the file containing the application's configuration
' dataset is to be determined.
'
'@/
Public Enum ConfigFileOptions
    ConfigFileOptionNone            ' The file is called settings.xml and is located in the application's settings folder.
    ConfigFileOptionFirstArg        ' The file's path and filename are specified in the first command line argument.
    ConfigFileOptionConfigSwitch    ' The file's path and filename are specified in the /config command line switch.
    ConfigFileOptionSettingsSwitch  ' The file's path and filename are specified in the /settings command line switch.
End Enum

''
' Used to set the foreground and background colours in the <code>Console</code>
' class.
'
' @param ConsoleBlack
'   The colour black.
' @param ConsoleBlue
'   The colour blue.
' @param ConsoleCyan
'   The colour cyan.
' @param ConsoleDarkBlue
'   The colour dark blue.
' @param ConsoleDarkCyan
'   The colour dark cyan.
' @param ConsoleDarkGray
'   The colour dark gray.
' @param ConsoleDarkGreen
'   The colour dark green.
' @param ConsoleDarkMagenta
'   The colour dark magenta
' @param ConsoleDarkRed
'   The colour dark red.
' @param ConsoleDarkYellow
'   The colour dark yellow.
' @param ConsoleGray
'   The colour gray.
' @param ConsoleGreen
'   The colour green.
' @param ConsoleMagenta
'   The colour magenta.
' @param ConsoleRed
'   The colour reg.
' @param ConsoleWhite
'   The colour white.
' @param ConsoleYellow
'   The colour white.
'@/
Public Enum ConsoleColors
    ConsoleBlack = 0
    ConsoleBlue = FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    ConsoleCyan = FOREGROUND_BLUE Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    ConsoleDarkBlue = FOREGROUND_BLUE
    ConsoleDarkCyan = FOREGROUND_BLUE Or FOREGROUND_GREEN
    ConsoleDarkGray = FOREGROUND_INTENSITY
    ConsoleDarkGreen = FOREGROUND_GREEN
    ConsoleDarkMagenta = FOREGROUND_RED Or FOREGROUND_BLUE
    ConsoleDarkRed = FOREGROUND_RED
    ConsoleDarkYellow = FOREGROUND_RED Or FOREGROUND_GREEN
    ConsoleGray = FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE
    ConsoleGreen = FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    ConsoleMagenta = FOREGROUND_RED Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    ConsoleRed = FOREGROUND_RED Or FOREGROUND_INTENSITY
    ConsoleWhite = FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    ConsoleYellow = FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
End Enum

''
' Used when creating a <code>SortedDictionary</code> object to indicate the type
' of Keys allowed.
'
' @param KeyTypeInteger
'   Keys must be integral values (ie a whole number).
' @param KeyTypeFloat
'   Keys must be floating point values.
' @param KeyTypeString
'   Keys must be strings. Case is ignored.
' @param KeyTypeCaseSensitiveString
'   Keys must be strings. Case is significant.
' @param KeyTypeDate
'   Keys must be dates.
' @param KeyTypeComparable
'   Keys must be objects that implement the <code>Comparable</code> interface.
'@/
Public Enum DictionaryKeyTypes
    KeyTypeInteger
    KeyTypeFloat
    KeyTypeString
    KeyTypeCaseSensitiveString
    KeyTypeDate
    KeyTypeComparable
End Enum

''
' Used to identify the units in which the expiry time for a timer
' is expressed.
'
' @param ExpiryTimeUnitDateTime
'   The expiry time is a date and time.
' @param ExpiryTimeUnitSeconds
'   The expiry time is an interval in seconds.
' @param ExpiryTimeUnitMilliseconds
'   The expiry time is an interval in milliseconds.
' @param ExpiryTimeUnitMinutes
'   The expiry time is an interval in minutes.
' @param ExpiryTimeUnitHours
'   The expiry time is an interval in hours.
' @param ExpiryTimeUnitDays
'   The expiry time is an interval in days.
'@/
Public Enum ExpiryTimeUnits
    ExpiryTimeUnitDateTime
    ExpiryTimeUnitSeconds
    ExpiryTimeUnitMilliseconds
    ExpiryTimeUnitMinutes
    ExpiryTimeUnitHours
    ExpiryTimeUnitDays
End Enum

''
' Used with the GetSpecialFolderPath global method to identify the folder required.
'
' @param FolderIdInternet
' @param FolderIdPrograms
' @param FolderIdMyDocuments
' @param FolderIdFavorites
' @param FolderIdStartup
' @param FolderIdRecent
' @param FolderIdSendTo
' @param FolderIdStartMenu
' @param FolderIdMyMusic
' @param FolderIdMyVideo
' @param FolderIdDesktop
' @param FolderIdNethood
' @param FolderIdFonts
' @param FolderIdTemplates
' @param FolderIdCommonStartmenu
' @param FolderIdCommonPrograms
' @param FolderIdCommonStartup
' @param FolderIdCommonDesktop
' @param FolderIdAppdata
' @param FolderIdPrinthood
' @param FolderIdLocalAppdata
' @param FolderIdCommonFavorites
' @param FolderIdInternetCache
' @param FolderIdCookies
' @param FolderIdHistory
' @param FolderIdCommonAppdata
' @param FolderIdWindows
' @param FolderIdSystem
' @param FolderIdProgramFiles
' @param FolderIdMyPictures
' @param FolderIdProfile
' @param FolderIdSystemX86
' @param FolderIdProgramFilesCommon
' @param FolderIdCommonTemplates
' @param FolderIdCommonDocuments
' @param FolderIdCommonAdmintools
' @param FolderIdAdmintools
' @param FolderIdConnections
' @param FolderIdCommonMusic
' @param FolderIdCommonPictures
' @param FolderIdCommonVideo
' @param FolderIdResources
' @param FolderIdResourcesLocalized
' @param FolderIdCommonOemLinks
' @param FolderIdCdBurnArea
Public Enum FolderIdentifiers
' These values are identical to the corresponding CSIDL_ constants used by Windows
' Shell API
'    FolderIdDesktop = &H0
    FolderIdInternet = &H1
    FolderIdPrograms = &H2
'    FolderIdControls = &H3
'    FolderIdPrinters = &H4
    FolderIdMyDocuments = &H5 ' maps to CSIDLIdPERSONAL
    FolderIdFavorites = &H6
    FolderIdStartup = &H7
    FolderIdRecent = &H8
    FolderIdSendTo = &H9
'    FolderIdBitBucket = &HA
    FolderIdStartMenu = &HB
'    FolderIdMyDocuments = &HC ' gives an error - use FolderIdPERSONAL instead
    FolderIdMyMusic = &HD
    FolderIdMyVideo = &HE
    FolderIdDesktop = &H10 ' maps to CSIDLDESKTOPDIRECTORY
'    FolderIdDrives = &H11
'    FolderIdNetwork = &H12
    FolderIdNethood = &H13
    FolderIdFonts = &H14
    FolderIdTemplates = &H15
    FolderIdCommonStartmenu = &H16
    FolderIdCommonPrograms = &H17
    FolderIdCommonStartup = &H18
    FolderIdCommonDesktop = &H19 ' maps to CSIDLCOMMONDESKTOPDIRECTORY
    FolderIdAppdata = &H1A
    FolderIdPrinthood = &H1B
    FolderIdLocalAppdata = &H1C
'    FolderIdAltstartup = &H1D
'    FolderIdCommonAltstartup = &H1E
    FolderIdCommonFavorites = &H1F
    FolderIdInternetCache = &H20
    FolderIdCookies = &H21
    FolderIdHistory = &H22
    FolderIdCommonAppdata = &H23
    FolderIdWindows = &H24
    FolderIdSystem = &H25
    FolderIdProgramFiles = &H26
    FolderIdMyPictures = &H27
    FolderIdProfile = &H28
    FolderIdSystemX86 = &H29
'    FolderIdProgramFilesX86 = &H2A
    FolderIdProgramFilesCommon = &H2B
'    FolderIdProgramFilesCommonX86 = &H2C
    FolderIdCommonTemplates = &H2D
    FolderIdCommonDocuments = &H2E
    FolderIdCommonAdmintools = &H2F
    FolderIdAdmintools = &H30
    FolderIdConnections = &H31
    FolderIdCommonMusic = &H35
    FolderIdCommonPictures = &H36
    FolderIdCommonVideo = &H37
    FolderIdResources = &H38
    FolderIdResourcesLocalized = &H39
    FolderIdCommonOemLinks = &H3A
    FolderIdCdBurnArea = &H3B
End Enum

''
' Generic error codes used in TradeWright products
'
' @param ErrArithmeticException
'   An exceptional arithmetic condition has occurred.
' @param ErrClassCastException
'   Attempt to cast an object to class of which it is not an instance.
' @param ErrIllegalArgumentException
'   Method has been passed an illegal or inappropriate argument.
' @param ErrIllegalStateException
'   Method has been invoked at an illegal or inappropriate time.
' @param ErrIndexOutOfBoundsException
'   An index of some sort is out of range.
' @param ErrNullPointerException
'   Attempt to use Nothing in a case where an object is required.
' @param ErrNumberFormatException
'   Attempt to convert a string to one of the numeric types, but the string does not have the appropriate format.
' @param ErrRuntimeException
'   An unspecified runtime error has occurred.
' @param ErrSecurityException
'   A security violation has occurred.
' @param ErrUnsupportedOperationException
'   The requested operation is not supported.
' @param ErrEndOfFile
'   The end of a file has been reached.
'@/
Public Enum ErrorCodes
    ErrArithmeticException = vbObjectError + 1024
    ErrClassCastException = vbObjectError + 1025
    ErrIllegalArgumentException = vbObjectError + 1026
    ErrIllegalStateException = vbObjectError + 1027
    ErrIndexOutOfBoundsException = vbObjectError + 1028
    ErrNullPointerException = vbObjectError + 1029
    ErrNumberFormatException = vbObjectError + 1030
    ErrRuntimeException = vbObjectError + 1031
    ErrSecurityException = vbObjectError + 1032
    ErrUnsupportedOperationException = vbObjectError + 1033
    ErrEndOfFile = vbObjectError + 1034
End Enum

''
' Used to indicate the type of a file.
'
' @param FileTypeUnknown
'   No information is available about the type of the file.
' @param FileTypeChar
'   The file is a character file, typically an LPT device or a console.
' @param FileTypeDisk
'   The file is a disk file.
' @param FileTypePipe
'   The file is either a named or anonymous pipe.
'@/
Public Enum FileTypes
    FileTypeUnknown = FILE_TYPE_UNKNOWN
    FileTypeChar = FILE_TYPE_CHAR
    FileTypeDisk = FILE_TYPE_DISK
    FileTypePipe = FILE_TYPE_PIPE
End Enum

''
'   The state space for an <code>IFuture</code> object.
'
' @param FutureStatePending
'   The <code>IFuture</code>'s value is not yet available but may become so later.
' @param FutureStateCompleted
'   The <code>IFuture</code>'s value is available.
' @param FutureStateCancelled
'   The <code>IFuture</code>'s value will never be available at the application's request.
' @param FutureStateErrored
'   An error has occurred during creation of the <code>IFuture</code>'s value.
'@/
Public Enum FutureStates
    FutureStatePending
    FutureStateCompleted
    FutureStateCancelled
    FutureStateErrored
End Enum

''
' Used to control the level of detail or importance of information that is
' logged.
'
' @param LogLevelNull
'   The <code>Logger</code> object will obtain its log level from its parent. This value
'   must only be used to set the <code>LogLevel</code> property of <code>Logger</code>
'   objects. It must not be used in logging methods.
' @param LogLevelUseDefault
'   The <code>Logger</code> object will use the value of the global <code>DefaultLogLevel</code>
'   property as its log level. This value must only be used to set the <code>LogLevel</code>
'   property of <code>Logger</code> objects. It must not be used in logging methods.
' @param LogLevelAll
'   The <code>Logger</code> object will log all information. This value
'   must only be used to set the <code>LogLevel</code> property of <code>Logger</code>
'   objects. It must not be used in logging methods.
' @param LogLevelHighDetail
'   Use this level for very detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to specialist support staff or developers, or is
'   very voluminous.
' @param LogLevelMediumDetail
'   Use this level for fairly detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to specialist support staff or developers, and is
'   not very voluminous.
' @param LogLevelDetail
'   Use this level for moderately detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to support staff or developers, and is
'   not very voluminous.
' @param LogLevelNormal
'   Use this level for general information that may be of interest to end users
'   or administrators, and is not very voluminous, but that does not generally
'   need to be brought to the user's attention.
' @param LogLevelInfo
'   Use this level for information that may be of interest to end users or administrators,
'   and that should be brought to the user's attention.
' @param LogLevelWarning
'   Use this level for information that may indicate a potential problem or non-critical failure.
' @param LogLevelSevere
'   Use this level for information regarding a critical failure.
' @param LogLevelNone
'   The <code>Logger</code> object will log no information. This value
'   must only be used to set the <code>LogLevel</code> property of <code>Logger</code>
'   objects. It must not be used in logging methods.
'@/
Public Enum LogLevels
    LogLevelAll = &H80000000
    LogLevelHighDetail = -3000
    LogLevelMediumDetail = -2000
    LogLevelDetail = -1000
    LogLevelNormal = 0
    LogLevelInfo = 1000
    LogLevelWarning = 2000
    LogLevelSevere = 3000
    LogLevelUseDefault = &H7FFFFFFD
    LogLevelNone = &H7FFFFFFE
    LogLevelNull = &H7FFFFFFF
End Enum

''
' Used by the StateTableBuilder class to represent generic special states.
'
' @param StateError
'   Indicates that a transition to the error state has occurred. The error state
'   usually indicates that an event has occurred that is not expected in the
'   current state.
'@/
Public Enum SpecialStates
    StateError = -1
End Enum

''
' Used with the StateTableBuilder class to represent special combinations of
' conditions.
'
' @param NoConditions
'   Represents all conditions unset.
' @param AllConditions
'   Represents all conditions set.
'@/
Public Enum SpecialConditions
    NoConditions = 0
    AllConditions = &HFFFFFFFF
End Enum

''
' Used by the StateTableBuilder class to represent generic special actions.
'
' @param NoAction
'   Indicates that no Action is to be taken.
'@/
Public Enum SpecialActions
    NoAction = -1
End Enum

''
' Used by the StateTableBuilder class to represent generic special stimuli.
'
' @param StimulusAll
'   Represents all stimuli.
'@/
Public Enum SpecialStimuli
    StimulusAll = -1
End Enum

''
'   A task's Priority governs the amount of processor time it receives.
'
' @param PriorityNormal
'   The default.
' @param PriorityLow
'   The task receives roughly half the processor time of tasks with
'   Priority <code>PriorityNormal</code>.
' @param PriorityHigh
'   The task receives roughly twice the processor time of tasks with
'   Priority <code>PriorityNormal</code>.
'@/
Public Enum TaskPriorities
    PriorityNormal
    PriorityLow
    PriorityHigh
End Enum

''
'   The state space for a <code>Task</code> object.
'
' @param TaskStatePending
'   The task is waiting to receive processor time because
'   the maxiumum number of tasks (defined in the <code>TaskConcurrency</code>
'   property) are already running.
' @param TaskStateRunnable
'   The task is currently getting processor time.
' @param TaskStateSuspended
'   The task has suspended itself.
' @param TaskStateCompleted
'   The task has completed its processing.
'@/
Public Enum TaskStates
    TaskStatePending
    TaskStateRunnable
    TaskStateSuspended
    TaskStateCompleted
End Enum

''
' The state space for a <code>TimerListItem</code> object.
'
' @param TimerListItemStatePending
'   Waiting for expiry to occur.
' @param TimerListItemStateExpired
'   Expiry has occurred.
' @param TimerListItemStateCancelled
'   Cancelled before expiry.
'@/
Public Enum TimerListItemStates
    TimerListItemStatePending
    TimerListItemStateExpired
    TimerListItemStateCancelled
End Enum

''
' These values identify different Units for measuring elapsed time.
'
' @param TimePeriodNone
'   The unit is not specified.
' @param TimePeriodSecond
'   The unit is in seconds.
' @param TimePeriodMinute
'   The unit is in minutes
' @param TimePeriodHour
'   The unit is in hours
' @param TimePeriodDay
'   The unit is in days.
' @param TimePeriodWeek
'   The unit is in weeks.
' @param TimePeriodMonth
'   The unit is in months.
' @param TimePeriodYear
'   The unit is in years.
' @param TimePeriodTickMovement
'   The unit is a number of price Ticks. The corresponding time is the
'   period during which price moves in a range of the specified number of Ticks.
' @param TimePeriodTickVolume
'   The unit is a trade Tick. The corresponding time is the
'   period during which accumulated Tick Volume increases by the specified number.
' @param TimePeriodVolume
'   The unit is a contract or share traded. The corresponding time is the
'   period during which accumulated Volume increases by the specified number.
'@/
Public Enum TimePeriodUnits
    TimePeriodNone = 0
    TimePeriodSecond = 1
    TimePeriodMinute = 2
    TimePeriodHour = 3
    TimePeriodDay = 4
    TimePeriodWeek = 5
    TimePeriodMonth = 6
'    TimePeriodLunarMonth = 7
    TimePeriodYear = 8
    TimePeriodTickMovement = 9     ' for constant momentum, point & figure, etc
    TimePeriodTickVolume = 10
    TimePeriodVolume = 11
End Enum

''
' The state space for a <code>IntervalTimer</code> object.
'
' @param TimerRunning
'   The timer is running.
' @param TimerStopped
'   The timer is stopped.
Public Enum TimerStates
    TimerRunning = 1
    TimerStopped
End Enum

''
' Timestamp format specifiers for use with the <code>FormatTimestamp</code>
' global method.
'
' @param TimestampTimeOnly
'   The formatted Timestamp is of the form <em>hhmmss.lll</em> where <em>lll</em>
'   is milliseconds.
' @param TimestampDateOnly
'   The formatted Timestamp is of the form <em>yyyymmdd</em>.
' @param TimestampDateAndTime
'   The formatted Timestamp is of the form <em>yyyymmddhhmmss.lll</em> where <em>lll</em>
'   is milliseconds.
' @param TimestampTimeOnlyISO8601
'   The formatted Timestamp is of the form <em>hh:mm:ss.lll</em> where <em>lll</em>
'   is milliseconds.
' @param TimestampDateOnlyISO8601
'   The formatted Timestamp is of the form <em>yyyy-mm-dd</em>.
' @param TimestampDateAndTimeISO8601
'   The formatted Timestamp is of the form <em>yyyy-mm-dd hh:mm:ss.lll</em> where <em>lll</em>
'   is milliseconds.
' @param TimestampTimeOnlyLocal
'   The formatted Timestamp is of the form <em>&lt;time&gt;.lll</em>, where &lt;time&gt; is
'   the time in the format defined in the Control Panel's Regional and Language Options,
'   and <em>lll</em> is milliseconds.
' @param TimestampDateOnlyLocal
'   The formatted Timestamp is of the form <em>&lt;date&gt;.lll</em>, where &lt;date&gt; is
'   the date in the format defined in the Control Panel's Regional and Language Options,
'   and <em>lll</em> is milliseconds.
' @param TimestampDateAndTimeLocal
'   The formatted Timestamp is of the form <em>&lt;date &amp; time&gt;.lll</em>, where
'   &lt;date &amp; time&gt; is the date and time in the format defined in the Control Panel's
'   Regional and Language Options, and <em>lll</em> is milliseconds.
' @param TimestampCustom
'   The formatted Timestamp is of the form <em>&lt;customformat&gt;.lll</em>
'   where <em>&lt;customformat&gt;</em> is the date and time in the supplied custom format
'   and <em>lll</em> is milliseconds.
' @param TimestampNoMillisecs
'   Used additively in conjunction with other values, prevents the milliseconds
'   value being included in the Timestamp.
'@/
Public Enum TimestampFormats
    TimestampTimeOnly = 0
    TimestampDateOnly = 1
    TimestampDateAndTime = 2
    TimestampCustom = 3
    TimestampTimeOnlyISO8601 = 4
    TimestampDateOnlyISO8601 = 5
    TimestampDateAndTimeISO8601 = 6
    TimestampTimeOnlyLocal = 7
    TimestampDateOnlyLocal = 8
    TimestampDateAndTimeLocal = 9
    TimestampNoMillisecs = &H40000000
' @param TimestampIncludeTimezone
'   Used additively in conjunction with other values, results in a timezone specifier
'   being appended. The format of the timezone specifier is +|-hh:mm.
'    TimestampIncludeTimezone = &H80000000
End Enum

''
' Common error codes raised by the Visual Basic runtime.
'
' @param VbErrInvalidProcedureCall
' @param VbErrOverflow
' @param VbErrSubscriptOutOfRange
' @param VbErrDivisionByZero
' @param VbErrTypeMismatch
' @param VbErrFileNotFound
' @param VbErrFileAlreadyOpen
' @param VbErrFileAlreadyExists
' @param VbErrDiskFull
' @param VbErrPermissionDenied
' @param VbErrPathNotFound
' @param VbErrInvalidObjectReference
' @param VbErrInvalidPropertyValue
' @param VbErrInvalidPropertyArrayIndex
' @param VbErrElementAlreadyExists
'@/
Public Enum VBErrorCodes
    VbErrInvalidProcedureCall = 5
    VbErrOverflow = 6
    VbErrSubscriptOutOfRange = 9
    VbErrDivisionByZero = 11
    VbErrTypeMismatch = 13
    VbErrFileNotFound = 53
    VbErrFileAlreadyOpen = 55
    VbErrFileAlreadyExists = 58
    VbErrDiskFull = 61
    VbErrPermissionDenied = 70
    VbErrPathNotFound = 76
    VbErrInvalidObjectReference = 91
    VbErrInvalidPropertyValue = 380
    VbErrInvalidPropertyArrayIndex = 381
    VbErrElementAlreadyExists = 457
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type CancelledEventData
    Source          As Object
End Type

''
' Conveys details of a change to the Source object.
'
' @param Source
'   The object that has changed
' @param ChangeType
'   Specifies the nature of the change. The values used are specific to
'   the Source object's type, and should be specfied in the relevant class
'   documentation.
'@/
Public Type ChangeEventData
    Source          As Object
    ChangeType      As Long
End Type

''
' Conveys details of a change to the Source object, which contains a collection of some
' sort.
'
' @param Source
'   The object that has changed.
' @param ChangeType
'   Specifies the nature of the change. The value used must be a member of the
'   <code>CollectionChangeTypes</code> enum. The relevant class documentation should
'   specify which members of the <code>CollectionChangeTypes</code> enum may
'   be used.
' @param AffectedItem
'   The item affected by the change. This has the value <code>Empty</code> where
'   the change type is <code>CollCollectionCleared</code> or
'   <code>CollOrderChanged</code>.
'@/
Public Type CollectionChangeEventData
    Source          As Object
    ChangeType      As CollectionChangeTypes
    AffectedItem    As Variant
End Type

''
' Used internally by the deferred Action mechanism.
'
' @param Action
' @param Data
'@/
Public Type DeferredActionEntry
    Action              As IDeferredAction
    Data                As Variant
End Type

''
' This type may be used by classes that implement the <code>Enumerable</code> interface.
'
'@/
Public Type EnumerationCookie
    Current             As Variant  ' Information that provides access to the current item in the enumeration
    Next                As Variant  ' Information that provides access to the Next item in the enumeration
End Type

''
' Conveys details of an error arising within the Source object.
'
' @param Source
'   The object within which the error arose.
' @param ErrorCode
'   An error code identifying the nature of the error.
' @param errorMessage
'   A description of the error.
' @param errorSource
'   Details of where the error occurred.
'@/
Public Type ErrorEventData
    Source          As Object
    ErrorCode       As Long
    ErrorMessage    As String
    ErrorSource     As String
End Type

Public Type FutureWaitCompletedEventData
    Source          As Object
    Future          As IFuture
    ContinuationData    As Variant
End Type

''
' Conveys details of a significant occurrence within the Source object.
'
' @param Source
'   The object within which the situation occurred.
' @param EventCode
'   A numeric code identifying the nature of the occurrence.
' @param EventMessage
'   A description of the occurrence.
'@/
Public Type NotificationEventData
    Source          As Object
    EventCode       As Long
    EventMessage    As String
End Type

''
' Indicates that the value of one of the Source object's properties has changed.
'
' @param Source
'   The object within which the situation occurred.
' @param propertyName
'   The name of the property whose value has changed.
'@/
Public Type PropertyChangedEventData
    Source          As Object
    PropertyName    As String
End Type

''
' Notifies a change of State within the Source object
'
' @param Source
'   The object whose State has changed.
' @param State
'   The new State of the Source object. The relevant class documentation
'   should specify what values may be set here.
'@/
Public Type StateChangeEventData
    Source          As Object
    State           As Long
End Type

''
' Used internally within the <code>StateEngine</code> and
' <code>StateTableBuilder</code> classes.
'
' @param State
' @param stimulus
' @param conditionsTrue
' @param conditionsFalse
' @param newState
' @param actions()
'@/
Public Type StateTableEntry
    State As Long
    Stimulus As Long
    ConditionsTrue As Long
    ConditionsFalse As Long
    NewState As Long
    Actions() As Long
End Type

''
' Contains details of a command line switch.
'
' @param Switch
'   The switch identifier.
' @param Value
'   The switch value.
'@/
Public Type SwitchEntry
    Switch      As String
    Value       As String
End Type

''
' Contains details of a completed task (including a task that has caused an
' error to be raised or that has been cancelled).
'
' @param Source
'   The <code>TaskController</code> object for the task that has completed processing.
' @param result
'   A value generated by the task.
' @param Cancelled
'   Set to <code>True</code> if the task was cancelled by a call to its
'   <code>TaskController</code>'s <code>CancelTask</code> method.
' @param pCookie
'   The pCookie value that was passed to the <code>StartTask</code> global method
'   when the task was started.
' @param TaskName
'   The completed task's name.
' @param ErrorNumber
'   If the task terminated due to an error, this contains the error number.
' @param ErrorMessage
'   If the task terminated due to an error, this contains the error message.
' @param ErrorSource
'   If the task terminated due to an error, this contains information about
'   where the error occurred.
'@/
Public Type TaskCompletionEventData
    Source          As Object
    Result          As Variant
    Cancelled       As Boolean
    Cookie          As Variant
    TaskName        As String
    ErrorNumber     As Long
    ErrorMessage    As String
    ErrorSource     As String
End Type

''
' Indicates a task's Progress.
'
' @param Source
'   The <code>TaskController</code> object for the task that is reporting Progress.
' @param Progress
'   The percentage of the task that has been completed.
' @param pCookie
'   The pCookie value that was passed to the <code>StartTask</code> global method
'   when the task was started.
' @param InterimResult
'   A value generated by the task.
'@/
Public Type TaskProgressEventData
    Source          As Object
    Progress        As Single
    Cookie          As Variant
    InterimResult  As Variant
End Type

''
' Conveys details of a significant occurrence within the Source task.
'
' @param Source
'   The <code>TaskController</code> object for the task that reported the occurrence.
' @param EventCode
'   A numeric code identifying the nature of the occurrence.
' @param EventMessage
'   A description of the occurrence.
' @param pCookie
'   The pCookie value that was passed to the <code>StartTask</code> global method
'   when the task was started.
'@/
Public Type TaskNotificationEventData
    Source          As Object
    EventCode       As Long
    EventMessage    As String
    Cookie          As Variant
End Type

''
' Notifies a change of State within the Source task.
'
' @param Source
'   The <code>TaskController</code> object for the task whose State has changed.
' @param State
'   The new State of the Source object. This is a value from the <code>TaskStates</code>
'   enum.
' @param pCookie
'   The pCookie value that was passed to the <code>StartTask</code> global method
'   when the task was started.
'@/
Public Type TaskStateChangeEventData
    Source          As Object
    State           As TaskStates
    Cookie          As Variant
End Type

''
' Contains summary information about a task.
'
' @param Name
'   The task's name.
' @param ClassName
'   The class name of the task object.
' @param Priority
'   The task's Priority.
' @param State
'   The task's State.
' @param StartTime
'   The time at which the task was started.
' @param LastRunTime
'   The time at which the task's Run method was last called.
' @param TotalCPUTime
'   The total time in milliseconds consumed in processing the task.
'@/
Public Type TaskSummary
    Name            As String
    ClassName       As String
    Priority        As String
    State           As String
    StartTime       As Date
    LastRunTime     As Date
    TotalCPUTime    As Long
End Type

''
' Notifies the expiry of the source timer.
'
' @param Source
'   The <code>IntervalTimer</code> object that has expired.
'@/
Public Type TimerExpiredEventData
    Source          As Object
    Data            As Variant
End Type

'@================================================================================
' External Function Declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

' There should be no member variables, because they are probably expected to be
' global. Although this class is GlobalMultiUse, that doesn't in fact mean that
' every client will get the same instance of it. It appears that clients in
' different modules each get their own instance of it.

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Sets the application name that is used to determine the application settings
' folder, to create default logfiles, and in various other contexts.
' @remarks
'   The default if this property is not specified is 'Application'.
' @param Value
'   The required application name. The value supplied must not contain any of the
'   following characters: &lt; &gt; : \ \ / "
'@/
Public Property Let ApplicationName( _
                ByVal Value As String)
Const ProcName As String = "ApplicationName"
On Error GoTo Err

gAssertArgument Value <> "", "Application Name cannot be blank"

gRegExp.Pattern = "^[^<>:\\/""|]+$"
gAssertArgument gRegExp.Test(Value), "Application Name contains invalid characters"

gApplicationName = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the application name that is used to determine the application settings
' folder, to create default logfiles, and in various other contexts.
' @return
'   The application name.
'@/
Public Property Get ApplicationName() As String
ApplicationName = gApplicationName
End Property

''
' Sets the application group name that is used to determine the application settings
' folder.
' @remarks
'   The default if this property is not specified is ''.
' @param Value
'   The required application group name. The value supplied must not contain any of the
'   following characters: &lt; &gt; : \ \ / "
'@/
Public Property Let ApplicationGroupName( _
                ByVal Value As String)
Const ProcName As String = "ApplicationGroupName"
On Error GoTo Err

gRegExp.Pattern = "^[^<>:\\/""|]*$"
gAssertArgument gRegExp.Test(Value), "Application Name contains invalid characters"

gApplicationGroupName = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the application group name that is used to determine the application settings
' folder.
' @return
'   The application group name.
'@/
Public Property Get ApplicationGroupName() As String
ApplicationGroupName = gApplicationGroupName
End Property

''
' Returns the path of the application settings folder.
'
' @remarks
' @return
'   The application settings folder path.
'@/
Public Property Get ApplicationSettingsFolder() As String
Const ProcName As String = "ApplicationSettingsFolder"

On Error GoTo Err

ApplicationSettingsFolder = gApplicationSettingsFolder

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the path and filename of the log file created by a call to the
' <code>SetupDefaultLogging</code> method.
' @return
'   The path and filename of the log file.
' @param pCommandLine
'   A string that may contain flags used to control the log filename and logging level.
'   For a Visual Basic 6 program, this value may be obtained using the <code>Command</code> function.
' @remarks
'   See the documentation for the <code>SetupDefaultLogging</code> method for information
'   about the permitted flag formats in the <code>pCommandLine</code> argument.
'@/
Public Property Get DefaultLogFileName(ByVal pCommandLine As String) As String
Const ProcName As String = "DefaultLogFileName"
Static logFileName As String

On Error GoTo Err

If logFileName = "" Then
    If gGetCommandLineParser(pCommandLine).Switch(SwitchLogFilename) Then logFileName = gGetCommandLineParser(pCommandLine).SwitchValue(SwitchLogFilename)

    If logFileName = "" Then
        logFileName = ApplicationSettingsFolder & "\" & gApplicationName & ".log"
    End If
End If
DefaultLogFileName = logFileName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
''
' Sets the default log level for <code>Logger</code> objects whose <code>LogLevel</code>
' is not explicitly set.
'
' @remarks
'   If this property is not set by the application, a value of <code>LogLevelNormal</code>
'   is assumed.
' @param Value
'   The required default logg level.
'@/
Public Property Let DefaultLogLevel( _
                ByVal Value As LogLevels)
Const ProcName As String = "DefaultLogLevel"
On Error GoTo Err

If Value = LogLevelAll Then
    ' this is allowed
Else
    gAssertArgument gIsLogLevelPermittedForApplication(Value), "This Value is not permitted in this context"
End If

gDefaultLogLevel = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
                
''
' Gets the default log level for <code>Logger</code> objects whose <code>LogLevel</code>
' is not explicitly set.
' @return
'   The current default logging level.
'@/
Public Property Get DefaultLogLevel() As LogLevels
DefaultLogLevel = gDefaultLogLevel
End Property

''
' Sets a value which determines the thread Priority during task execution.
'
' @remarks
'   By default, the value is <code>False</code>.
'
' @param Value
'   If <code>True</code>, the thread Priority is set to one point below the normal
'   value while a task is being processed. This helps to ensure that the system
'   remains responsive while a lengthy CPU-intensive task is being processed.
'
'   If <code>False</code>, the thread Priority remains at the normal value while
'   a task is being processed.
' @see
'
'@/
Public Property Let RunTasksAtLowerThreadPriority( _
                ByVal Value As Boolean)
Const ProcName As String = "RunTasksAtLowerThreadPriority"
On Error GoTo Err

GTasks.gTaskManager.RunTasksAtLowerThreadPriority = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Sets a value which determines the thread Priority during task execution.
'
' @remarks
'   By default, the value is <code>False</code>.
'
' @param Value
'   If <code>True</code>, the thread Priority is set to one point below the normal
'   value while a task is being processed. This helps to ensure that the system
'   remains responsive while a lengthy CPU-intensive task is being processed.
'
'   If <code>False</code>, the thread Priority remains at the normal value while
'   a task is being processed.
' @see
'
'@/
Public Property Get RunTasksAtLowerThreadPriority() As Boolean
RunTasksAtLowerThreadPriority = GTasks.gTaskManager.RunTasksAtLowerThreadPriority
End Property

''
' Sets the maximum number of tasks that can be processed concurrently.
'
' @param Value
'   The maximum number of tasks that can be processed concurrently.
'   The minimum permitted value is 2.
' @remarks
'   When the number of tasks defined by this property are already running, starting
'   another task results in it being added to a pending tasks queue. When one of the
'   running tasks completes or is suspended, the first task is removed from the
'   pending queue and added to the running tasks list.
'
'   By default, the task concurrency property is set to a value that depends on the
'   computer's CPU speed. In most applications it should not be necessary to set
'   this property.
'@/
Public Property Let TaskConcurrency( _
                ByVal Value As Long)
Const ProcName As String = "TaskConcurrency"

On Error GoTo Err

GTasks.gTaskManager.TaskConcurrency = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Gets the maximum number of tasks that can be processed 'concurrently'.
'
' @return
'   The maximum number of tasks that can be processed 'concurrently'.
'@/
Public Property Get TaskConcurrency() As Long
TaskConcurrency = GTasks.gTaskManager.TaskConcurrency
End Property

' NB: this property is currently hidden
Public Property Get TaskManager() As TaskManager
Attribute TaskManager.VB_MemberFlags = "40"
Set TaskManager = gTaskManager
End Property

''
' Sets the time in milliseconds allowed for each quantum of processing during task
' execution.
'
' @param Value
'   The time in milliseconds allowed for each quantum of processing during task
'   execution. The minimum permitted value is 1.
' @remarks
'
'@/
Public Property Let TaskQuantumMillisecs( _
                ByVal Value As Long)
Const ProcName As String = "TaskQuantumMillisecs"

On Error GoTo Err

GTasks.gTaskManager.QuantumMillisecs = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Gets the time in milliseconds allowed for each quantum of processing during task
' execution.
'
' @returns
'   The time in milliseconds allowed for each quantum of processing during task
'   execution.
'@/
Public Property Get TaskQuantumMillisecs() As Long
TaskQuantumMillisecs = GTasks.gTaskManager.QuantumMillisecs
End Property

''
' Sets the time interval in seconds at which task summary information is logged.
'
' @param Value
'   The time in seconds between successive task summary logging. A value of zero means that
'   there is no such logging.
' @remarks
'   Task summary information has a log infotype of diag.twutilities.tasks, and is logged at Detail
'   level.
'
'@/
Public Property Let TaskSummaryLoggingIntervalSecs( _
                ByVal Value As Long)
On Error GoTo Err

GTasks.gTaskManager.TaskSummaryLoggingIntervalSecs = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TaskSummaryLoggingIntervalSecs", pModuleName:=ModuleName, pFailpoint:=""
End Property

''
' Indicates whether tracing is enabled for the null trace type.
'
' @remarks
'   Using this method avoids the need for the application to maintain a reference to
'   the null tracer object.
'@/
Public Property Get TracingEnabled() As Boolean
TracingEnabled = gGetTracer("").Enabled
End Property

Public Property Get UnhandledErrorHandler() As UnhandledErrorHandler
Set UnhandledErrorHandler = gUnhandledErrorHandler
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AdjustColorIntensity(ByVal pColor As Long, ByVal pFactor As Double) As Long
AdjustColorIntensity = gAdjustColorIntensity(pColor, pFactor)
End Function

''
' Specifies a condition that should be true. If not, an error is raised.
'
' @remarks
' @param pCondition
'   The condition that should be true.
' @param pMessage
'   The error message to be raised if the condition is not true.
' @param pException
'   The error number to be raised if the condition is not true. This defaults to
'   <code>ErrorCodes.ErrIllegalStateException</code>.
'@/
Public Sub Assert( _
                ByVal pCondition As Boolean, _
                Optional ByVal pMessage As String, _
                Optional ByVal pException As Long = ErrorCodes.ErrIllegalStateException)
gAssert pCondition, pMessage, pException
End Sub

''
' Specifies a condition that should be true. If not, an <code>ErrorCodes.ErrIllegalArgumentException</code> error is raised.
'
' @remarks
' @param pCondition
'   The condition that should be true.
' @param pMessage
'   The error message to be raised if the condition is not true.
'@/
Public Sub AssertArgument( _
                ByVal pCondition As Boolean, _
                Optional ByVal pMessage As String)
gAssert pCondition, pMessage, ErrorCodes.ErrIllegalArgumentException
End Sub

''
' Converts a byte array to a hexadecimal string.
'
' @remarks
' @return
'   the hexadecimal string corresponding to the supplied
'   byte array
' @param pArray
'   the byte array to be converted
'@/
Public Function BytesToHexString(pArray() As Byte) As String
Dim i As Long

BytesToHexString = gBytesToHexString(pArray)
End Function

''
' Converts a date in the computer's timezone to another timezone.
'
' @return
'   The specified date converted to the timezone represented by the
'   <code>tz</code> argument.
' @param pLocalDate
'   A date in the computer's timezone.
' @param pTimezone
'   The timezone to which the date is to be converted.
'@/
Public Function ConvertDateLocalToTz( _
                ByVal pLocalDate As Date, _
                ByVal pTimezone As TimeZone) As Date
Const ProcName As String = "ConvertDateLocalToTz"

On Error GoTo Err

ConvertDateLocalToTz = pTimezone.ConvertDateLocalToTz(pLocalDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in the computer's local timezone to UTC.
'
' @return
'   The specified date converted to UTC.
' @param pLocalDate
'   A local date.
'@/
Public Function ConvertDateLocalToUTC( _
                ByVal pLocalDate As Date) As Date
Const ProcName As String = "ConvertDateLocalToUTC"

On Error GoTo Err

ConvertDateLocalToUTC = gLocalToUtc(pLocalDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in another timezone to the computer's local timezone.
'
' @return
'   The specified date converted to the computer's local timezone.
' @param tzDate
'   A date in the timezone represented by the <code>tz</code> argument.
' @param tz
'   The timezone from which the date is to be converted.
'@/
Public Function ConvertDateTzToLocal( _
                ByVal tzDate As Date, _
                ByVal tz As TimeZone) As Date
Const ProcName As String = "ConvertDateTzToLocal"

On Error GoTo Err

ConvertDateTzToLocal = tz.ConvertDateTzToLocal(tzDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in one timezone to another timezone.
'
' @return
'   The specified date converted to the required timezone.
' @param pDate
'   A date in the timezone represented by the <code>tz1</code> argument.
' @param tz1
'   The timezone from which the date is to be converted.
' @param tz2
'   The timezone to which the date is to be converted.
'@/
Public Function ConvertDateTzToTz( _
                ByVal pDate As Date, _
                ByVal tz1 As TimeZone, _
                ByVal tz2 As TimeZone) As Date
Const ProcName As String = "ConvertDateTzToTz"

On Error GoTo Err

ConvertDateTzToTz = tz1.ConvertDateTzToTz(pDate, tz2)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in another timezone to UTC.
'
' @return
'   The specified date converted to UTC.
' @param LocalDate
'   A date in the timezone represented by the <code>tz</code> argument.
' @param tz
'   The timezone from which the date is to be converted.
'@/
Public Function ConvertDateTzToUTC( _
                ByVal localDate As Date, _
                ByVal tz As TimeZone) As Date
Const ProcName As String = "ConvertDateTzToUTC"

On Error GoTo Err

ConvertDateTzToUTC = tz.ConvertDateTzToUTC(localDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a UTC date to the computer's local timezone.
'
' @return
'   The specified date converted to local time.
' @param utcDate
'   A UTC date.
'@/
Public Function ConvertDateUTCToLocal( _
                ByVal utcDate As Date) As Date
Const ProcName As String = "ConvertDateUTCToLocal"

On Error GoTo Err

ConvertDateUTCToLocal = gUtcToLocal(utcDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a UTC date to another timezone.
'
' @return
'   The specified date converted to the timezone represented by the
'   <code>tz</code> argument.
' @param utcDate
'   A UTC date.
' @param tz
'   The timezone to which the date is to be converted.
'@/
Public Function ConvertDateUTCToTZ( _
                ByVal utcDate As Date, _
                ByVal tz As TimeZone) As Date
Const ProcName As String = "ConvertDateUTCToTZ"

On Error GoTo Err

ConvertDateUTCToTZ = tz.ConvertDateUTCToTZ(utcDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>BasicLogFormatter</code> object.
'
' @return
'   A <code>BasicLogFormatter</code> object.
' @param timestampFormat
'   Specifies the Timestamp format to be used when formatting log records.
' @param includeInfoType
'   Specifies whether the log infotype is to be included in the formatted log record.
' @param includeTimestamp
'   Specifies whether the log event's timestamp is to be included in the formatted log record.
' @param includeLogLevel
'   Specifies whether the log event's log level is to be included in the formatted log record.
'@/
Public Function CreateBasicLogFormatter( _
                Optional ByVal timestampFormat As TimestampFormats = TimestampDateAndTimeISO8601, _
                Optional ByVal includeInfoType As Boolean, _
                Optional ByVal includeTimestamp As Boolean = True, _
                Optional ByVal includeLogLevel As Boolean = True) As ILogFormatter
Const ProcName As String = "CreateBasicLogFormatter"
On Error GoTo Err

Dim formatter As BasicLogFormatter
Set formatter = New BasicLogFormatter
formatter.Initialise timestampFormat, includeInfoType, includeTimestamp, includeLogLevel

Set CreateBasicLogFormatter = formatter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBoxedValue(ByVal pValue As Variant) As BoxedValue
Const ProcName As String = "CreateBoxedValue"
On Error GoTo Err

Set CreateBoxedValue = New BoxedValue
CreateBoxedValue.Initialise pValue

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>CommandLineParser</code> object.
'
' @return
'   The required <code>CommandLineParser</code> object.
' @param commandLine
'   The command line arguments to be parsed. For a Visual Basic 6 program,
'   this value may be obtained using the <code>Command</code> function.
' @param separator
'   A single character used as the separator between command line arguments.
'@/
Public Function CreateCommandLineParser( _
                ByVal commandLine As String, _
                Optional ByVal separator As String = " ") As CommandLineParser
Const ProcName As String = "CreateCommandLineParser"
On Error GoTo Err

gAssertArgument Len(separator) = 1, "Separator must be a single character"

Set CreateCommandLineParser = New CommandLineParser
CreateCommandLineParser.Initialise commandLine, separator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateConfigurationStore( _
                ByVal pConfigStoreProvider As IConfigStoreProvider, _
                ByVal pFilename As String) As ConfigurationStore
Const ProcName As String = "CreateConfigurationStore"

On Error GoTo Err

Set CreateConfigurationStore = gCreateConfigurationStore(pConfigStoreProvider, pFilename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>ConsoleLogListener</code> object.
'
' @return
'   The required <code>ConsoleLogListener</code> object.
' @param formatter
'   A <code>LogFormatter</code> object that is to be used to format the log records. If
'   not supplied, a <code>BasicLogFormatter</code> object is used, and log infotype is not
'   included in the formatted log records.
' @param timestampFormat
'   Specifies the Timestamp format to be used when formatting log records. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeInfoType
'   Specifies whether the log infotype is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeTimestamp
'   Specifies whether the log event's timestamp is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeLogLevel
'   Specifies whether the log event's log level is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
'@/
Public Function CreateConsoleLogListener( _
                Optional ByVal formatter As ILogFormatter, _
                Optional ByVal timestampFormat As TimestampFormats = TimestampDateAndTimeISO8601, _
                Optional ByVal includeInfoType As Boolean, _
                Optional ByVal includeTimestamp As Boolean = True, _
                Optional ByVal includeLogLevel As Boolean = True) As ILogListener
Const ProcName As String = "CreateConsoleLogListener"
On Error GoTo Err

Dim lListener As ConsoleLogListener

Set lListener = New ConsoleLogListener
lListener.Initialise formatter, timestampFormat, includeInfoType, includeTimestamp, includeLogLevel
Set CreateConsoleLogListener = lListener

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates an <code>Enumerator</code> object for interating the list members
' of an object that implements the <code>Enumerable</code> interface.
'
' @return
'   An <code>Enumerator</code> object.
' @param pTarget
'   The object for which an <code>Enumerator</code> is required.
' @param pCookie
'   The initial value for the pCookie used to keep track of the iteration.
'@/
Public Function CreateEnumerator( _
                ByVal pTarget As IEnumerable, _
                Optional ByVal pCookie As Variant) As Enumerator
Set CreateEnumerator = New Enumerator
CreateEnumerator.Initialise pTarget, pCookie
End Function

''
' Creates a <code>FileLogListener</code> object.
'
' @return
'   The required <code>FileLogListener</code> object.
' @param pFilename
'   The file to which the log records should be written.
' @param formatter
'   A <code>LogFormatter</code> object that is to be used to format the log records. If
'   not supplied, a <code>BasicLogFormatter</code> object is used, and log infotype is not
'   included in the formatted log records.
' @param pOverwrite
'   Indicates whether the existing contents of the log file should be overwritten.
' @param pCreateBackup
'   If <code>True</code>, then if the specified file already exists and the
'   <code>overwrite</code> argument is <code>False</code>, then the existing
'   file is renamed by having ".bak" appended to its filename.
' @param pUnicode
'   Indicates whether the characters written to the file are pUnicode-encoded.
' @param timestampFormat
'   Specifies the Timestamp format to be used when formatting log records. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeInfoType
'   Specifies whether the log infotype is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeTimestamp
'   Specifies whether the log event's timestamp is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
' @param includeLogLevel
'   Specifies whether the log event's log level is to be included in the formatted log record. This is only
'   relevant if the<code>formatter</code> argument is not supplied.
'@/
Public Function CreateFileLogListener( _
                ByVal pFilename As String, _
                Optional ByVal formatter As ILogFormatter, _
                Optional ByVal pOverwrite As Boolean, _
                Optional ByVal pCreateBackup As Boolean, _
                Optional ByVal pUnicode As Boolean, _
                Optional ByVal timestampFormat As TimestampFormats = TimestampDateAndTimeISO8601, _
                Optional ByVal includeInfoType As Boolean, _
                Optional ByVal includeTimestamp As Boolean = True, _
                Optional ByVal includeLogLevel As Boolean = True) As ILogListener
Const ProcName As String = "CreateFileLogListener"
On Error GoTo Err

Dim lListener As FileLogListener
Set lListener = New FileLogListener
lListener.Initialise pFilename, formatter, pOverwrite, pCreateBackup, pUnicode, timestampFormat, includeInfoType, includeTimestamp, includeLogLevel
Set CreateFileLogListener = lListener

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateFormattingLogger( _
                ByVal pInfoType As String, _
                ByVal pProjectName As String) As FormattingLogger
Const ProcName As String = "CreateFormattingLogger"
On Error GoTo Err

Set CreateFormattingLogger = gCreateFormattingLogger(pInfoType, pProjectName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateFuture( _
                ByVal pObject As Object, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "CreateFutureFromTask"
On Error GoTo Err

Dim lBuilder As New FutureBuilder
lBuilder.Value = pObject
lBuilder.Cookie = pCookie
lBuilder.Complete

Set CreateFuture = lBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateFutureFromTask( _
                ByVal pTaskController As TaskController) As IFuture
Const ProcName As String = "CreateFutureFromTask"
On Error GoTo Err

Dim lBuilder As New TaskFutureBuilder
lBuilder.Initialise pTaskController
Set CreateFutureFromTask = lBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates an <code>IntervalTimer</code> object.
'
' @param firstExpiryTime
'   Specifies when the first <code>TimerExpired</code> event is to occur. This can be
'   specified as an elapsed time from the current time, or as a date and time, as
'   indicated by the <code>firstExpiryUnits</code> argument.
' @param firstExpiryUnits
'   The units in which the<code>firstExpiryTime</code> argument is specified.
' @param repeatIntervalMillisecs
'   If not zero, <code>TimerExpired</code> events are fired at that interval
'   until the timer is stopped.
' @param useRandomIntervals
'   If <code>True</code>, the periods between <code>TimerExpired</code> events are
'   a random Length up to the Length specified.
' @param pData
'   Any required data. This is returned when the timer fires.
'@/
Public Function CreateIntervalTimer( _
                ByVal firstExpiryTime As Variant, _
                Optional ByVal firstExpiryUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds, _
                Optional ByVal repeatIntervalMillisecs As Long, _
                Optional ByVal useRandomIntervals As Boolean, _
                Optional ByVal pData As Variant) As IntervalTimer
Const ProcName As String = "CreateIntervalTimer"
On Error GoTo Err

Set CreateIntervalTimer = gCreateIntervalTimer(firstExpiryTime, firstExpiryUnits, repeatIntervalMillisecs, useRandomIntervals, pData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>Parameters</code> object with parameter values parsed from
' a string.
' @remarks
'   The parameter names and values are supplied in a <code>String</code>
'   with the following syntax:
'   <pre>
'   <tt>
'   <p>paramString: param [sep param]...</p>
'
'   <p>param: name delim value</p>
'
'   <p>name: string</p>
'
'   <p>value: string</p>
'   </tt>
'   </pre>
'   where
'   <pre>
'   <p><tt>sep</tt> is a single character used as a separator between parameters.
'   The default separator is a semi-colon, but an alternative can be specified
'   in the <code>parameterSeparator</code> argument.</p>
'
'   <p><tt>delim</tt> is a single character used to delimit the name
'   The default delimiter is =, but an alternative can be specified
'   in the <code>nameDelimiter</code> argument.</p>
'   </pre>
'
'   <p>If the parameter name or value needs to contain a separator or delimiter
'   character, this can be achieved by preceding it with an escape character.
'   The default escape character is \, but an alternative can be specified
'   in the <code>escapeCharacter</code> argument. Note that to include the
'   escape character itself in the value or name, it must be repeated. The
'   effect of the escape character is to include the following character in
'   the parameter name or value verbatim.</p>
'
'   <p>For each specified parameter, if a parameter with this name is
'   already included, its value is updated; otherwise a new parameter
'   is added to the list.</p>
' <h3>Examples</h3>
'   An example parameter string, using the default separator and delimiter
'   characters:
'   <pre>
'   <tt>Server=Delphi;Port=7496;ClientId=432</tt>
'   </blockquote>
'   An example incuding the escape character in the value:
'   <pre>
'   <tt>SettingsFile=C:\\MyProg\\myprog.ini</tt>
'   </pre>
'   An example using the escape character in the value in a unusual (and pointless!) way:
'   <pre>
'   <tt>My parameter=\a\b\c\d\e\f</tt>
'   </pre>
'   Note that in this case, the value would be interpreted as abcdef.
' @return
'   A <code>Parameters</code> object initialised with the parameters specified in
'   the <code>paramString</code> argument.
' @param paramString
'   The parameters to be set, using the syntex described under Remarks.
' @param NameDelimiter
'   A single character used to delimit the parameter name.
' @param parameterSeparator
'   A single character used to separate parameters.
' @param escapeCharacter
'   A single character that causes the following character to be included
'   verbatim in the parameter value in spite of any special meaning it may normally have.
'@/
Public Function CreateParametersFromString( _
                ByVal paramString As String, _
                Optional ByVal NameDelimiter As String = "=", _
                Optional ByVal parameterSeparator As String = ";", _
                Optional ByVal escapeCharacter As String = "\") As Parameters
Const ProcName As String = "CreateParametersFromString"
On Error GoTo Err

Set CreateParametersFromString = gCreateParametersFromString(paramString, NameDelimiter, parameterSeparator, escapeCharacter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>ParameterStringParser</code> object for parsing parameters from
' a string.
' @remarks
'   The parameter names and values are supplied in a <code>String</code>
'   with the following syntax:
'   <pre>
'   <tt>
'   <p>paramString: param [sep param]...</p>
'
'   <p>param: name delim value</p>
'
'   <p>name: string</p>
'
'   <p>value: string</p>
'   </tt>
'   </pre>
'   where
'   <pre>
'   <p><tt>sep</tt> is a single character used as a separator between parameters.
'   The default separator is a semi-colon, but an alternative can be specified
'   in the <code>parameterSeparator</code> argument.</p>
'
'   <p><tt>delim</tt> is a single character used to delimit the name
'   The default delimiter is =, but an alternative can be specified
'   in the <code>nameDelimiter</code> argument.</p>
'   </pre>
'
'   <p>If the parameter name or value needs to contain a separator or delimiter
'   character, this can be achieved by preceding it with an escape character.
'   The default escape character is \, but an alternative can be specified
'   in the <code>escapeCharacter</code> argument. Note that to include the
'   escape character itself in the value or name, it must be repeated. The
'   effect of the escape character is to include the following character in
'   the parameter name or value verbatim.</p>
'
'   <p>For each specified parameter, if a parameter with this name is
'   already included, its value is updated; otherwise a new parameter
'   is added to the list.</p>
' <h3>Examples</h3>
'   An example parameter string, using the default separator and delimiter
'   characters:
'   <pre>
'   <tt>Server=Delphi;Port=7496;ClientId=432</tt>
'   </pre>
'   An example incuding the escape character in the value:
'   <pre>
'   <tt>SettingsFile=C:\\MyProg\\myprog.ini</tt>
'   </pre>
'   An example using the escape character in the value in a unusual (and pointless!) way:
'   <pre>
'   <tt>My parameter=\a\b\c\d\e\f</tt>
'   </pre>
'   Note that in this case, the value would be interpreted as abcdef.
' @return
'   A <code>ParameterStringParser</code> object to parse the parameters specified in
'   the <code>paramString</code> argument.
' @param paramString
'   The parameters to be parsed, using the syntex described under Remarks.
' @param NameDelimiter
'   A single character used to delimit the parameter name.
' @param parameterSeparator
'   A single character used to separate parameters.
' @param escapeCharacter
'   A single character that causes the following character to be included
'   verbatim in the parameter value in spite of any special meaning it may normally have.
'@/
Public Function CreateParameterStringParser( _
                ByVal paramString As String, _
                Optional ByVal NameDelimiter As String = "=", _
                Optional ByVal parameterSeparator As String = ";", _
                Optional ByVal escapeCharacter As String = "\") As ParameterStringParser
Const ProcName As String = "CreateParameterStringParser"

On Error GoTo Err

Set CreateParameterStringParser = gCreateParameterStringParser(paramString, _
                                                            NameDelimiter, _
                                                            parameterSeparator, _
                                                            escapeCharacter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a simulated <code>Clock</code> object running in a specified timezone.
'
' @return
'   The required <code>Clock</code> object.
' @param pRate
'   The factor by which time is accelerated in this clock.
' @param pTimezonename
'   The standard name of the required timezone.
'@/
Public Function CreateSimulatedClock( _
                ByVal pRate As Single, _
                ByVal pTimezonename As String) As Clock
Const ProcName As String = "CreateSimulatedClock"
On Error GoTo Err

gAssertArgument pRate >= 0, "Clock rate cannot be less than zero"

Set CreateSimulatedClock = GClock.gCreateSimulatedClock(pRate, pTimezonename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>SortedDictionary</code> object using Keys of a specified type.
'
' @return
'   The required <code>SortedDictionary</code> object.
' @param keyType
'   Specifies the type of Keys allowed. This must be a member of the
'   DictionaryKeyTypes enumeration.
' @param allowDuplicates
'   Indicates whether entries with duplicate Keys are allowed.
'@/
Public Function CreateSortedDictionary( _
                Optional ByVal keyType As DictionaryKeyTypes = KeyTypeString, _
                Optional ByVal allowDuplicates As Boolean = False) As SortedDictionary
Const ProcName As String = "CreateSortedDictionary"
On Error GoTo Err

Set CreateSortedDictionary = New SortedDictionary
CreateSortedDictionary.Initialise keyType, allowDuplicates

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>StateEngine</code> object with a State table constructed by a
' <code>StateTableBuilder</code> object.
'
' @remarks
'   The State table must have been fully built (ie the <code>StateTableBuilder</code>
'   object's <code>stateTableComplete</code> method must have been called) before
'   calling this method.
' @return
'   A <code>StateEngine</code> object.
' @param tableBuilder
'   The <code>StateTableBuilder</code> object that was used to construct the State
'   table.
' @param initialState
'   An application-defined State identifier for the starting State.
'@/
Public Function CreateStateEngine( _
                ByVal TableBuilder As StateTableBuilder, _
                ByVal initialState As Long) As StateEngine
Const ProcName As String = "CreateStateEngine"
On Error GoTo Err

Set CreateStateEngine = New StateEngine
CreateStateEngine.Initialise TableBuilder, initialState

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a new <code>StringBuilder</code> object.
'
' @return
'   The required <code>StringBuilder</code> object.
' @param initialValue
'   A value to be set in the <code>StringBuilder</code> object as creation time,
'   as if the <code>append</code> method were called explicitly.
' @param initialCapacity
'   The number of characters that the <code>StringBuilder</code> object can
'   hold initially. Setting this argument unnecessarily large or small can
'   lead to some inefficiency, however its value is not critical.
' @param maximumCapacity
'   The maximum size that the string held in this <code>StringBuilder</code>
'   is allowed to grow to. It is only necessary to use this argument when there is
'   a specific need to limit the size of the string. Its value has no effect on
'   efficiency.
'@/
Public Function CreateStringBuilder( _
                Optional ByVal initialvalue As String = "", _
                Optional ByVal initialCapacity As Long = 128, _
                Optional ByVal maximumCapacity As Long = &HFFFFFFF) As StringBuilder
Const ProcName As String = "CreateStringBuilder"
On Error GoTo Err

Set CreateStringBuilder = New StringBuilder
CreateStringBuilder.Initialise initialCapacity, maximumCapacity, initialvalue

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>WeakReference</code> object targetting a specified object.
'
' @return
'   The required <code>WeakReference</code> object .
' @param Target
'   The object for which a weak reference is required.
'@/
Public Function CreateWeakReference( _
                ByVal Target As Object) As WeakReference
Set CreateWeakReference = gCreateWeakReference(Target)
End Function

Public Function CreateWriteableTextFile( _
                ByVal pFilename As String, _
                Optional ByVal pOverwrite As Boolean, _
                Optional ByVal pCreateBackup As Boolean, _
                Optional ByVal pUnicode As Boolean, _
                Optional ByVal pIncrementFilenameIfInUse As Boolean) As TextStream
Const ProcName As String = "CreateWriteableTextFile"
On Error GoTo Err

Set CreateWriteableTextFile = gCreateWriteableTextFile(pFilename, pOverwrite, pCreateBackup, pUnicode, pIncrementFilenameIfInUse)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates an <code>XMLConfigProvider</code> object.
'
' @remarks
'   Note that the return type is <code>ConfigStoreProvider</code>. This is because
'   <code>XMLConfigProvider</code> implements the <code>ConfigStoreProvider</code>
'   interface, and it is via that interface that it must be accessed.
' @return
'   The required <code>XMLConfigProvider</code> object.
' @param pApplicationName
'   The name of the application for which this provider will manage configuration
'   data.
' @param pApplicationVersion
'   The version of the application for which this provider will manage configuration
'   data.
'@/
Public Function CreateXMLConfigurationProvider( _
                Optional ByVal pApplicationName As String, _
                Optional ByVal pApplicationVersion As String) As IConfigStoreProvider
Const ProcName As String = "CreateXMLConfigurationProvider"
On Error GoTo Err

Set CreateXMLConfigurationProvider = gCreateXMLConfigurationProvider(pApplicationName, pApplicationVersion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Queues a <code>DeferredAction</code> object for later processing.
'
' @param pAction
'   The Action to be queued for processing. This object must implement the
'   <code>DeferredAction</code> interface.
' @param pData
'   A value that will be passed to the <code>run</code> method of the
'   <code>DeferredAction</code> object.
' @param pDeferTime
'   Specifies when the Action should be processed. If this has a value of 0 or is
'   not specified, the Action will be processed as soon as possible. Note however
'   that the Action will never be processed before the processing of the current
'   event has completed.
' @param pDeferTimeUnits
'   The units in which the<code>deferTime</code> argument is specified.
' @param pMayYield
'   This parameter is no longer used, and its value is ignored.
'@/
Public Sub DeferAction( _
                ByVal pAction As IDeferredAction, _
                Optional ByVal pData As Variant, _
                Optional ByVal pDeferTime As Variant, _
                Optional ByVal pDeferTimeUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds, _
                Optional ByVal pMayYield As Boolean)
Const ProcName As String = "DeferAction"
On Error GoTo Err

DeferredActionManager.DeferAction pAction, pData, pDeferTime, pDeferTimeUnits

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
''
' Disables tracing for the specified trace type.
'
' @param pTraceType
'   The trace type for which tracing is to be disabled. See the documentation
' for the<code>Tracer</code> class for further information about trace types.
'@/
Public Sub DisableTracing( _
                ByVal pTraceType As String)
Const ProcName As String = "DisableTracing"
On Error GoTo Err

gDisableTracing pTraceType

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Parses a string representation of a double and returns the Double value.
'
' Note that the string representation must use a period as the decimal point separator.
' Unlike <code>CDbl</code>, <code>DoubleFromString</code> does not observe
' locale conventions.
'
' @return
'   The <code>Double</code> value parsed.
'
' @param v
'   The string to be parsed
'
'@/
Public Function DoubleFromString(ByVal v As String) As Double
Const ProcName As String = "DoubleFromString"
On Error GoTo Err

DoubleFromString = gDoubleFromString(v)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns string representation of the supplied Double value.
'
' Note that the string representation uses a period as the decimal
' point separator. Unlike <code>CStr</code>, <code>DoubleToString</code>
' does not observe locale conventions.
'
' @return
'   The converted <code>Double</code>.
'
' @param d
'   The value to be converted to a string.
'
'@/
Public Function DoubleToString(ByVal d As Double) As String
Const ProcName As String = "DoubleToString"
On Error GoTo Err

DoubleToString = gDoubleToString(d)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Enables tracing for the specified trace type.
'
' @param pTraceType
'   The trace type for which tracing is to be enabled. See the documentation
' for the<code>Tracer</code> class for further information about trace types.
'@/
Public Sub EnableTracing( _
                ByVal pTraceType As String)
Const ProcName As String = "EnableTracing"
On Error GoTo Err

gEnableTracing pTraceType

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the specified value into a JSON string.
'
' JSON is JavaScript Object Notation - see http://www.json.org and
' http://www.ietf.org/rfc/rfc4627.txt?number=4627. JSON is a lightweight, text-based,
' language-independent data interchange format.
'
' @remarks
'   The supplied value must be either an array variant, or an object variant.
'
'   Object variants are encoded by taking the following rules in order:
'
'   <ol>
'   <li>If the object implements the <code>JSONable</code> interface, then the
'   encoding is delegated to the object (ie the result of the <code>ToJSON</code> method
'   is returned.</li>
'   <li>If the object is of type <code>Dictionary</code> or <code>SortedDictionary</code>,
'   then the keys and the values of the items contained in the object are encoded.</li>
'   <li>If the object is of type <code>Collection</code> or implements the
'   <code>Enumerable</code> interface, then the encoding is as if the object were
'   an array variant, since the keys (if any) used when storing the items in the
'   object cannot be ascertained.</li>
'   <li>If the object has a default property, then a single name/value pair is encoded
'   with name 'DefaultProp' and value the encoding of the default property's value.</li>
'   </ol>
'
' @return
'   The JSON string representing the supplied value.
' @param Value
'   The value to be encoded.
' @see
'
'@/
Public Function EncodeJSON( _
                ByVal Value As Variant) As String
Const ProcName As String = "EncodeJSON"
On Error GoTo Err

Dim sb As New StringBuilder
sb.Initialise
GJSON.gEncode Value, sb
EncodeJSON = sb.ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Terminates the current process and all of its threads.
'
' @remarks
'   This method should only be called when a fatal error has occurred that cannot be
'   recovered from. The process is ended abruptly.
'
'   Note that applications that use this component should call this method rather
'   than the Visual Basic <code>End</code> statement, which sometimes fails to terminate
'   the process when this component is in use.
' @param pExitCode
'   Specifies the exit code for the process.
' @see
'
'@/
Public Sub EndProcess(Optional ByVal pExitCode As Long = 1)
gEndProcess pExitCode
End Sub

''
' Formats timestamps as a string.
'
' A milliseconds value is appended in the form ".lll" (eg "20070809161723.238"),
' unless the <code>formatOption</code>
' argument is set to <code>TimestampFormats.TimestampDateOnly</code> or
' <code>TimestampFormats.TimestampDateOnlyISO8601</code> or includes the value
' <code>TimestampFormats.TimestampNoMillisecs</code>.
'
' @return
'   The formatted Timestamp.
' @param Timestamp
'   The Timestamp to be formatted.
' @param formatOption
'   Specifies the format to be used. The <code>TimestampFormats.TimestampNoMillisecs</code>
'   value can be included along with another <code>TimestampFormats</code>
'   values if milliseconds are not to be included.
' @param formatString
'   If the <code>formatOption</code> argument is set to
'   <code>TimestampFormats.TimestampCustom</code>, then this argument contains
'   the custom format string to be used for formatting the non-millseconds part
'   of the Timestamp, to which the miliseconds part is appended (unless
'   <code>TimestampNoMillisecs</code> is specified in the <code>formatOption</code>
'   argument).
'@/
Public Function FormatTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal formatOption As TimestampFormats = TimestampFormats.TimestampDateAndTime, _
                Optional ByVal formatString As String = "yyyymmddhhnnss") As String
Const ProcName As String = "FormatTimestamp"
On Error GoTo Err

FormatTimestamp = gFormatTimestamp(Timestamp, formatOption, formatString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Generates a GUID (Globally Unique IDentifier). This is a 128-bit value which is
' (to all intents and purposes) guaranteed to be unique for all time.
'
' @remarks
'   Since Visual Basic has no way to represent a 128-bit value directly, the
'   GUID is returned in a code>GUIDStruct</code> user-defined type, which
'   contains a four-element array of <code>Long</code>s to hold the value.
' @return
'   The GUID value.
'@/
Public Function GenerateGUID() As GUIDStruct
GenerateGUID = gGenerateGUID
End Function

''
' Returns the string representation of a generated GUID (Globally Unique IDentifier -
' this is a 128-bit value which is (to all intents and purposes) guaranteed to be
' unique for all time).
'
' @remarks
'   The returned value is formatted according to the widely-accepted convention
'   for GUIDs, for example:
'   <pre>
'   <tt>c200e360-38c5-11ce-ae62-08002b2b79ef</tt>
'   </pre>
'
' @return
'   The GUID string value.
'@/
Public Function GenerateGUIDString() As String
GenerateGUIDString = gGenerateGUIDString
End Function

''
' Generates a pseudo-random <code>Long</code>number.
'
' @remarks
'   Repeated calls to this method will eventually generate duplicate values.
'   However it is unlikely (though possible) that a short sequence of calls will generate
'   duplicates. (In other words there is no guarantee that the returned value is unique.)
' @return
'   A pseudo-random <code>Long</code>number.
'@/
Public Function GenerateID() As Long
GenerateID = gGenerateID
End Function

''
' Generates a string representation of a pseudo-random <code>Long</code>number.
'
' @remarks
'   Repeated calls to this method will eventually generate duplicate values.
'   However it is unlikely (though possible) that a short sequence of calls will generate
'   duplicates.
' @return
'   A string representation of a generated pseudo-random <code>Long</code>number.
'@/
Public Function GenerateTextID() As String
GenerateTextID = gGenerateTextID
End Function

''
' Returns the standard names of all timezones known to the system.
'
' @remarks
'   The timezone names are those held by Windows in the registry.
' @return
'   An array containing the known timezone names.
'@/
Public Function GetAvailableTimeZoneNames() As String()
Const ProcName As String = "GetAvailableTimeZoneNames"
On Error GoTo Err

GetAvailableTimeZoneNames = GTimeZone.gGetAvailableTimeZoneNames

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the global <code>Clock</code> object for a specified timezone.
'
' @return
'   The required <code>Clock</code> object.
' @param pTimezonename
'   The standard name of the required timezone.
'@/
Public Function GetClock( _
                Optional ByVal pTimezonename As String) As Clock
Const ProcName As String = "GetClock"
On Error GoTo Err

Set GetClock = GClock.gGetClock(pTimezonename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the command line used to start the current program.
'
' @remarks
'   This method differs from VB6's <code>Command</code> method in the following ways:
'
'   <ul>
'   <li>the value returned includes the program's executable filename at the start, before the arguments</li>
'   <li>it works when called from within a library, whereas the <code>Command</code> method returns
'   an empty string</li>
'   <li>when called from a program running within the VB6 IDE, it returns the command line for
'   VB6 itself, rather than for the program under test</li>
'   </ul>
' @return
'   The command line used to start the current program.
'@/
Public Function GetCommandLine() As String
Const ProcName As String = "GetCommandLine"
On Error GoTo Err

GetCommandLine = gGetCommandLine

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the console object for this program.
'
' @remarks
'   If the program was started with a console, a <code>Console</code> object
'   giving access to it is returned.
'
'   If the program was started without a console (ie it is a GUI program), a
'   console is allocated and a <code>Console</code> object giving access to it is returned.
'
' @return
'   The <code>Console</code> object.
'@/
Public Function GetConsole() As Console
Const ProcName As String = "GetConsole"
On Error GoTo Err

Set GetConsole = gGetConsole

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the amount of CPU time in seconds consumed by the current process.
'
' @remarks
'
' @return
'   The total CPU time in seconds consumed by the current process.
'@/
'
Public Function GetCurrentProcessCpuTime() As Single
GetCurrentProcessCpuTime = gGetCurrentProcessCpuTime
End Function

''
' Returns a <code>ConfigurationStore</code> object giving access to the application's
' configuration dataset (optionally creating it).
'
' @remarks
'   Uses the <code>XMLConfigProvider</code> class to provide the configuration dataset, which
'   is therefore in the form of an XML file.
'
'   Returns <code>Nothing</code> if a configuration dataset does not already exist and the
'   <code>pCreate</code> argument is set to <code>False</code>.
'
'   If the value of the <code>pOptions</code> argument is <code>ConfigFileOptionsNone</code>,
'   or no file name is specified in the indicated location, then it is
'   assumed that the configuration dataset is located in the application's settings folder
'   and is called settings.xml.
' @return
'   A <code>ConfigurationStore</code> object, or <code>Nothing</code> if no valid configuration
'   dataset already exists and the <code>pCreate</code> argument is set to <code>False</code>.
' @param pCommandLine
'   The command line arguments that may need to be parsed. For a Visual Basic 6 program,
'   this value may be obtained using the <code>Command</code> function.
' @param pConfigFileVersion
'   If the configuration dataset already exists, a check is made that its version matches this value. If
'   not, a <code>ErrorCodes.ErrIllegalStateException</code> is raised.
'
'   If the configuration dataset has to be created, this value is stored in it.
' @param pCreate
'   If set to <code>True</code>, the configuration file is created. Note that any existing
'   configuration file will be overwritten.
' @param pOptions
'   Indicates where the location of the file containing the configuration dataset is specified.
' @param pIgnoreInvalid
'   If <code>True</code>, indicates that no <code>ErrorCodes.ErrIllegalStateException</code>
'   is to be raised if there is an existing configuration file that is no longer valid (eg it
'   has the wrong file version).
'@/
Public Function GetDefaultConfigurationStore( _
                ByVal pCommandLine As String, _
                ByVal pConfigFileVersion As String, _
                Optional ByVal pCreate As Boolean = False, _
                Optional ByVal pOptions As ConfigFileOptions = ConfigFileOptionFirstArg, _
                Optional ByVal pIgnoreInvalid As Boolean = False) As ConfigurationStore
Const ProcName As String = "GetDefaultConfigurationStore"
On Error GoTo Err

Set GetDefaultConfigurationStore = gGetDefaultConfigurationStore( _
                                                pCommandLine, _
                                                pConfigFileVersion, _
                                                pCreate, _
                                                pIgnoreInvalid, _
                                                pOptions)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the global realtime <code>TimerList</code> object.
'
' @return
'   The global realtime <code>TimerList</code> object.
'@/
Public Function GetGlobalTimerList() As TimerList
Const ProcName As String = "GetGlobalTimerList"
On Error GoTo Err

Set GetGlobalTimerList = gGetTimerList

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a simulated-time <code>TimerList</code> object.
'
' @return
'   A simulated-time <code>TimerList</code> object.
'
' @remarks
'   Calls using the same pClock parameter return the same <code>TimerList</code> object.
'@/
Public Function GetSimulatedTimerList( _
                ByVal pClock As Clock) As TimerList
Const ProcName As String = "GetSimulatedTimerList"
On Error GoTo Err

gAssertArgument pClock.IsSimulated, "The supplied Clock must be simulated"

Set GetSimulatedTimerList = gGetSimulatedTimerList(pClock)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the <code>Logger</code> object for a particular info type.
'
' @remarks
'   Information types form a hierarchical namespace, components of the name
'   being separated by a period character (ie '.'). The root of the tree is
'   an empty string. Information type names are not case sensitive.
'
' @return
'   The <code>Logger</code> object for the specified info type.
' @param InfoType
'   The info type for which a logger is required.
' @see
'
'@/
Public Function GetLogger( _
                ByVal InfoType As String) As Logger
Const ProcName As String = "GetLogger"
On Error GoTo Err

Set GetLogger = gLogManager.GetLogger(InfoType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a key that uniquely identifies the specified object.
'
' @param pObject
'   The object whose key is to be returned.
' @return
'   The specified object's key.
'@/
Public Function GetObjectKey(ByVal pObject As Object) As String
GetObjectKey = gGetObjectKey(pObject)
End Function

''
' Returns the ProgId for a COM object.
'
' @param obj
'   The object whose ProgId is to be returned.
' @return
'   The ProgId.
'@/
Public Function GetProgIdFromObject( _
                ByVal obj As Object) As String
Dim lTLI As New tli.TLIApplication
Dim ti As TypeInfo
Dim ii As InterfaceInfo

Const ProcName As String = "GetProgIdFromObject"
On Error GoTo Err

Set ti = lTLI.ClassInfoFromObject(obj)
Set ii = ti.DefaultInterface

GetProgIdFromObject = ii.Parent.Name & "." & _
                    IIf(Left$(ii.Name, 1) = "_", Replace(ii.Name, "_", "", 1, 1), ii.Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the full path to one of the Windows special folders.
'
' @remarks
'   This function is used to obtain the path for folders whose actual
'   location may vary from system to system or from user to user.
' @return
'   The full path to the required folder.
' @param folderId
'   Specifies which special folder is sought.
'@/
Public Function GetSpecialFolderPath( _
                ByVal folderId As FolderIdentifiers) As String
Const ProcName As String = "GetSpecialFolderPath"
On Error GoTo Err

GetSpecialFolderPath = gGetSpecialFolderPath(folderId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a <code>TimePeriod</code> object representing the specified period of time.
'
' @return
'   The required <code>TimePeriod</code> object.
' @param Length
'   The number of units of time in the period.
' @param Units
'   The unit of time in which the <code>Length</code> argument is expressed.
'@/
Public Function GetTimePeriod( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits) As TimePeriod
Const ProcName As String = "GetTimePeriod"
On Error GoTo Err

Set GetTimePeriod = gGetTimePeriod(Length, Units)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns an accurate Timestamp in the computer's local timezone.
'
' @return
'   The Timestamp.
'@/
Public Function GetTimestamp() As Date
Const ProcName As String = "GetTimestamp"
On Error GoTo Err

GetTimestamp = gGetTimestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns an accurate UTC Timestamp.
'
' @return
'   The Timestamp.
'@/
Public Function GetTimestampUTC() As Date
Const ProcName As String = "GetTimestampUTC"
On Error GoTo Err

GetTimestampUTC = gGetTimestampUtc

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the global <code>TimeZone</code> object for the specified timezone.
'
' @return
'   The requried <code>TimeZone</code> object.
' @param pTimezonename
'   The standard name for the required timezone.
'@/
Public Function GetTimeZone( _
                Optional ByVal pTimezonename As String) As TimeZone
Const ProcName As String = "GetTimeZone"
On Error GoTo Err

Set GetTimeZone = gGetTimeZone(pTimezonename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the <code>Tracer</code> object for the specified trace type.
'
' @param pTraceType
'   The trace type for which a <code>Tracer</code> object is required. See the documentation
' for the<code>Tracer</code> class for further information about trace types.
' @return
'   The required <code>Tracer</code> object.
'@/
Public Function GetTracer( _
                ByVal pTraceType As String) As Tracer
Const ProcName As String = "GetTracer"
On Error GoTo Err

Set GetTracer = gGetTracer(pTraceType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a string representation of the supplied <code>GUIDStruct</code>.
'
' @remarks
'   The returned value is formatted according to the widely-accepted convention
'   for GUIDs, for example:
'   <pre>
'   <tt>c200e360-38c5-11ce-ae62-08002b2b79ef</tt>
'   </pre>
' @return
'   The GUID string value.
' @param pGUID
'   The value whose string representation is required.
'@/
Public Function GUIDToString(ByRef pGUID As GUIDStruct) As String
GUIDToString = gGUIDToString(pGUID)
End Function

Public Sub HandleWin32Error()
gHandleWin32Error
End Sub

''
' Provides a common method for handling unexpected runtime errors.
'
' @remarks
'   This method can be called in the error handler of every procedure to deal with
'   unexpected errors (note that the term 'procedure' covers both <code>Sub</code>s,
'   <code>Function</code>s and <code>Property</code>s).
'
'   It has two purposes:
'   <ul>
'   <li>To log details of the error to the logging system.</li>
'   <li>To re-raise the error with additional Source information.</li>
'   </ul>
'
'   If the method is requested to re-raise the error, then the error with its additional
'   Source information will be caught by a higher-level procedure in the calling stack. If
'   the error is eventually caught by a procedure whose error handler calls this method
'   requesting logging, then the logged information will include details of each lower-level
'   procedure that called this method and requested re-raising. Thus if every procedure in
'   the call stack invokes this method in its error handler, the logged information will
'   effectively give the call stack that lead to the error.
'
'   It is recommended that every Private procedure should call the method in its error
'   handler requesting that the error be re-raised but not logged.
'
'   It is recommended that every procedure that may be called from other modules (ie
'   <code>Public</code> and <code>Friend</code> procedures) should call the method in its
'   error handler requesting that the error be re-raised but not logged. This enables
'   calling procedures to handle the error in whatever way they wish (including possibly
'   a further call to this method). It is acceptable to request that the error be logged
'   in case no higher-level procedure logs it, though this may result in duplicated log
'   information.
'
'   It is recommended that event handlers should NOT call this method. They
'   should instead call the <code>UnhandledErrorHandler</code>'s <code>Notify</code> method in their
'   error handler. This gives the application the opportunity to handle the error.
'
'   Note that when running your program in the IDE, your <code>Err</code> object is not shared
'   by the TWUtilities component, and it is therefore advisable to explictly pass the
'   <code>pErrorNumber</code>, <code>pErrorDesc</code> and <code>pErrorSource</code> arguments
'   to ensure that incorrect values are not assumed.
' @param pProcedureName
'   The name of the procedure in which the error was caught.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pFailpoint
'   A value that indicates the region of the calling procedure where the error occurred.
' @param pReRaise
'   If set to <code>True</code>, this method re-raises the error with additional
'   Source information, and the method does not return. If set to <code>False</code>,
'   the method returns normally.
' @param pLog
'   If set to <code>True</code>, the method logs details of the error to the logging system
'   with an infotype of "error" at loglevel <code>LogLevelSevere</code>.
' @param pErrorNumber
'   The error number. If this argument is not included in the call to this method, the current
'   value of <code>Err.Number</code> is used.
' @param pErrorDesc
'   The error description. If this argument is not included in the call to this method, the current
'   value of <code>Err.Description</code> is used.
' @param pErrorSource
'   The error Source. If this argument is not included in the call to this method, the current
'   value of <code>Err.Source</code> is used.
'@/
Public Sub HandleUnexpectedError( _
                ByVal pProcedureName As String, _
                ByVal pProjectName As String, _
                ByVal pModuleName As String, _
                Optional ByVal pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByVal pErrorDesc As String, _
                Optional ByVal pErrorSource As String)
gHandleUnexpectedError pProcedureName, pModuleName, pFailpoint, pReRaise, pLog, pErrorNumber, pErrorDesc, pErrorSource, pProjectName
End Sub

''
' Converts a string containing hexadecimal characters into the equivalent
' array of bytes.
'
' @return
'   A byte array.
' @param inString
'   description
'@/
Public Function HexStringToBytes(inString As String) As Byte()
Const ProcName As String = "HexStringToBytes"
On Error GoTo Err

Dim i As Long
Dim inAr() As Byte
Dim outAr() As Byte

HexStringToBytes = gHexStringToBytes(inString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Initialises the component.
'
' @remarks
'   It is advisable, but not necessary, for applications to call this
'   method at startup (typically in the <code>Form Initialize</code> event).
'   The reason is that initialising the component can take a significant
'   fraction of a second, during which no events can be processed. If the
'   initialisation is not done explicitly, then it will be done implicitly
'   the first time that any of several facilities is used, which may cause
'   some inconvenience if rapid processing of events is essential to the application.
'@/
Public Sub InitialiseTWUtilities()
Globals.gInitialise
End Sub

Public Sub InitialiseCommonControls()
Dim s As INITCOMMONCONTROLSSTRUCT
s.dwSize = Len(s)
s.dwICC = ICC_WIN95_CLASSES + ICC_DATE_CLASSES + ICC_USEREX_CLASSES + ICC_COOL_CLASSES
InitCommonControlsEx s
End Sub

''
' Converts an integer array to a hexadecimal string.
'
' @return
'   the hexadecimal string corresponding to the supplied
'   integer array
' @param inAr
'   the integer array to be converted
'@/
Public Function IntegersToHexString( _
                ByRef inAr() As Integer) As String
Const ProcName As String = "IntegersToHexString"
On Error GoTo Err

IntegersToHexString = gIntegersToHexString(inAr)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


''
' Indicates whether the supplied value is a whole number, optionally within
' a specified range.
'
' @return
'   <code>True</code> if the value is a whole number within the relevant range.
' @param Value
'   The value to be checked.
' @param minValue
'   The lowest acceptable value. Must not be greater than the <code>maxValue</code>
'   argument.
' @param maxValue
'   The highest acceptable value. Must not be less than the <code>minValue</code>
'   argument.
'@/
Public Function IsInteger( _
                ByVal Value As Variant, _
                Optional ByVal minValue As Long = &H80000000, _
                Optional ByVal maxValue As Long = &H7FFFFFFF) As Boolean
Const ProcName As String = "IsInteger"
On Error GoTo Err

IsInteger = gIsInteger(Value, minValue, maxValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsRelativeConfigPath( _
                ByVal configPath As String) As Boolean
Const ProcName As String = "IsRelativeConfigPath"
On Error GoTo Err

Dim cp As ConfigurationPath

Set cp = gGetConfigPath(configPath)
IsRelativeConfigPath = cp.IsRelative

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidConfigAttributePath( _
                ByVal configPath As String) As Boolean
Const ProcName As String = "IsValidConfigAttributePath"
On Error GoTo Err

Dim cp As ConfigurationPath

Set cp = gGetConfigPath(configPath)
IsValidConfigAttributePath = cp.IsValidAttributePath

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidConfigValuePath( _
                ByVal configPath As String) As Boolean
Const ProcName As String = "IsValidConfigValuePath"
On Error GoTo Err

Dim cp As ConfigurationPath

Set cp = gGetConfigPath(configPath)
IsValidConfigValuePath = cp.IsValidValuePath

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidConfigSectionPath( _
                ByVal configPath As String) As Boolean
Const ProcName As String = "IsValidConfigSectionPath"
On Error GoTo Err

Dim cp As ConfigurationPath

Set cp = gGetConfigPath(configPath)
IsValidConfigSectionPath = cp.IsValidSectionPath

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates an <StdFont</code> object from properties stored in
' a <code>ConfigurationSection</code> object.
' @return
'   An <code>StdFont</code> object initialised with the properties stored in
'   the <code>config</code> argument.
' @param config
'   A <code>ConfigurationSection</code> object containing the font properties.
'@/
Public Function LoadFontFromConfig( _
                ByVal Config As ConfigurationSection) As StdFont
Const ProcName As String = "LoadFontFromConfig"
On Error GoTo Err

Dim aFont As New StdFont
aFont.Bold = Config.GetSetting(ConfigSettingFontBold, aFont.Bold)
aFont.Italic = Config.GetSetting(ConfigSettingFontItalic, aFont.Italic)
aFont.Name = Config.GetSetting(ConfigSettingFontName, aFont.Name)
aFont.Size = Config.GetSetting(ConfigSettingFontSize, aFont.Size)
aFont.Strikethrough = Config.GetSetting(ConfigSettingFontStrikethrough, aFont.Strikethrough)
aFont.Underline = Config.GetSetting(ConfigSettingFontUnderline, aFont.Underline)
Set LoadFontFromConfig = aFont

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>Parameters</code> object from parameter values stored in
' a <code>ConfigurationSection</code> object.
' @return
'   A <code>Parameters</code> object initialised with the parameters stored in
'   the <code>config</code> argument.
' @param config
'   A <code>ConfigurationSection</code> object containing the parameters.
'@/
Public Function LoadParametersFromConfig( _
                ByVal Config As ConfigurationSection) As Parameters
Const ProcName As String = "LoadParametersFromConfig"
On Error GoTo Err

Set LoadParametersFromConfig = New Parameters
LoadParametersFromConfig.LoadFromConfig Config

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


''
' Loads a file of XML configuration Data into an <code>XMLConfigProvider</code> object.
'
' @remarks
'   Note that the return type is <code>ConfigStoreProvider</code>. This is because
'   <code>XMLConfigProvider</code> implements the <code>ConfigStoreProvider</code>
'   interface, and it is via that interface that it must be accessed.
' @return
'   The required <code>XMLConfigProvider</code> object.
' @param filePath
'   The path and filename of the file containing XML configuration
'   Data.
'@/
Public Function LoadConfigProviderFromXMLFile( _
                ByVal filePath As String) As IConfigStoreProvider
Const ProcName As String = "LoadConfigProviderFromXMLFile"
On Error GoTo Err

Set LoadConfigProviderFromXMLFile = gLoadConfigProviderFromXMLFile(filePath)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LogMessage(ByVal pMsg As String, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "LogMessage"
On Error GoTo Err

If Not gLogLogger.IsLoggable(pLogLevel) Then Exit Sub
gLogLogger.Log pLogLevel, pMsg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Translates a string value into the corresponding <code>LogLevels</code> enumeration
' type value.
'
' @return
'   The <code>LogLevels</code> value corresponding to the supplied string.
' @param Value
'   The string to be converted to a <code>LogLevels</code> value. This should be one of
'   the following (note that this is not case-sensitive):
'
'   "A", "ALL"
'   "D", "DETAIL"
'   "H", "HIGH", "HIGHDETAIL", "HIGH DETAIL"
'   "I", "INFO"
'   "M", "MEDIUM", "MEDIUMDETAIL", "MEDIUM DETAIL"
'   "0", "NONE"
'   "N", "NORMAL"
'   "-", "NULL"
'   "S", "SEVERE"
'   "W", "WARNING"
'   "U", "DEFAULT"
'
'   Any other supplied value raise an <code>ErrorCodes.ErrIllegalArgumentException</code>
'   error.
'@/
Public Function LogLevelFromString( _
                ByVal Value As String) As LogLevels
LogLevelFromString = gLogLevelFromString(Value)
End Function

''
' Translates a <code>LogLevels</code> value to the corresponding single character.
'
' @return
'   The single-character string corresponding to the supplied <code>LogLevels</code> value.
' @param Value
'   The <code>LogLevels</code> value to be converted to a string.
'
'@/
Public Function LogLevelToShortString( _
                ByVal Value As LogLevels) As String
LogLevelToShortString = gLogLevelToShortString(Value)
End Function

''
' Translates a <code>LogLevels</code> value to the corresponding string.
'
' @return
'   The string corresponding to the supplied <code>LogLevels</code> value.
' @param Value
'   The <code>LogLevels</code> value to be converted to a string.
'
'@/
Public Function LogLevelToString( _
                ByVal Value As LogLevels) As String
LogLevelToString = gLogLevelToString(Value)
End Function

''
' Converts an array of <code>Long</code>s to a hexadecimal string.
'
' @return
'   the hexadecimal string corresponding to the supplied
'   long array
' @param inAr
'   the long array to be converted
'@/
Public Function LongsToHexString( _
                inAr() As Integer) As String
Const ProcName As String = "LongsToHexString"
On Error GoTo Err

LongsToHexString = gLongsToHexString(inAr)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
''
' The latest date that can be held in a <code>Date</code> variable.
'
'@/
Public Function MaxDate() As Date
MaxDate = MaxDateValue
End Function

''
' The largest value that can be held in a <code>Double</code> variable.
'
'@/
Public Function MaxDouble() As Double
MaxDouble = MaxDoubleValue
End Function

''
' The earliest date that can be held in a <code>Date</code> variable.
'
'@/
Public Function MinDate() As Date
MinDate = MinDateValue
End Function

''
' The smallest value that can be held in a <code>Double</code> variable.
'
'@/
Public Function MinDouble() As Double
MinDouble = MinDoubleValue
End Function

Public Function MonthStartDate( _
                ByVal pDate As Date) As Date
Const ProcName As String = "MonthStartDate"
On Error GoTo Err

MonthStartDate = gCalcMonthStartDate(pDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MonthStartDateFromMonthNumber( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "MonthStartDateFromMonthNumber"
On Error GoTo Err

MonthStartDateFromMonthNumber = gCalcMonthStartDateFromMonthNumber(monthNumber, baseDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NormalizeColor(ByVal pColor As Long) As Long
NormalizeColor = gNormalizeColor(pColor)
End Function

''
' Returns the supplied value if it is non-null, or the default value if it is null.
'
' @return
'   If the <code>Value</code> argument is not null, it is returned. If it is null,
'   the <code>defaultValue</code> argument is returned.
' @param Value
'   The value to be tested for nullity.
' @param defaultValue
'   The value to be returned if the <code>Value</code> argument is null.
'@/
Public Function Nz( _
                ByVal Value As Variant, _
                Optional ByVal defaultValue As Variant) As Variant
If IsNull(Value) Then
    Nz = defaultValue
Else
    Nz = Value
End If
End Function

''
' Parses a JSON string and returns a corresponding value.
'
' JSON is JavaScript Object Notation - see http://www.json.org and
' http://www.ietf.org/rfc/rfc4627.txt?number=4627. JSON is a lightweight, text-based,
' language-independent data interchange format.
'
' @remarks
'   The type of the value returned in the <code>result</code> argument is governed by
'   the following rules:
'   <ul>
'       <li>If the JSON string encodes an array, the value is a variant array of variants.</li>
'       <li>If the first name/value pair in the JSON string has name '$ProgId',
'           the value is a variant containing an instance of the class whose progId
'           is the corresponging value, provided the following conditions are satisfied:
'           <ul>
'               <li>the class has its <code>Instancing</code> property set to
'               <code>Public</code>;</li>
'               <li>for each name/value pair after the first in the JSON encoding, the class
'               has a writable property of the same name that accepts the encoded value.</li>
'           </ul>
'       </li>
'       <li>Otherwise, the value is a <code>Dictionary</code> object whose entries contain
'           the name/value pairs (each entry is keyed by the name and contains the value as
'           the item).</li>
'   </ul>
' @param inputString
'   The JSON to be parsed.
' @param result
'   A <code>Variant</code> value built from the parsed JSON string.
' @see
'
'@/
Public Sub ParseJSON( _
                ByVal inputString As String, _
                ByRef Result As Variant)
Const ProcName As String = "ParseJSON"
On Error GoTo Err

GJSON.gParse inputString, Result

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Replaces a method of a class with another procedure.
' @param pObjectPointer
'   The address of an instance of the class whose method is to be redirected. This
'   can be obtained using the <code>ObjPtr</code> function.
' @param pvTableIndex
'   The index in the vTable of the method to be redirected.
' @param pNewAddress
'   The address of the procedure that is to replace the original method. This can be obtained
'   using the AddressOf operator.
' @remarks
'   The method is redirected for all existing and future instances of the class, not
'   just the instance whose address is passed in the <code>pObjectPointer</code> argument.
'
'   Note that where a class implements an interface, and it is a method of this interface
'   that is to be redirected, the value of <code>pObjectPointer</code> must be
'   obtained by applying the <code>ObjPtr</code> function to a variable of the correct
'   interface type containing a reference to the object.
'
'   The procedure whose address is passed in the <code>pNewAddress</code> argument must
'   be defined in a standard code module (not a class module), must have a first
'   argument that is of type <code>Object</code>, and the remaining arguments and
'   the return value (if any) must be identical to those of the method being redirected. When
'   the procedure is invoked, the first argument will be a reference to the particular instance
'   of the class on which the method is being called.
'@/
Public Function RedirectMethod( _
                ByVal pObjectPointer As Long, _
                ByVal pvTableIndex As Long, _
                ByVal pNewAddress As Long) As Long
RedirectMethod = gRedirectMethod(pObjectPointer, pvTableIndex, pNewAddress)
End Function

''
' Saves an <code>StdFont</code> object to a <code>ConfigurationSection</code> object.
' @param aFont
'   The <code>StdFont</code> object to be saved.
' @param config
'   A <code>ConfigurationSection</code> object where the font properties are to be stored.
' @remarks
'   The <code>config</code> argument should refer to a <code>ConfigurationSection</code>
'   object that will only be used to store the font properties. Using it to store other
'   configuration settings may lead to unpredictable results.
'@/
Public Sub SaveFontToConfig( _
                ByVal aFont As StdFont, _
                ByVal Config As ConfigurationSection)
Const ProcName As String = "SaveFontToConfig"
On Error GoTo Err

Config.SetSetting ConfigSettingFontName, aFont.Name
Config.SetSetting ConfigSettingFontBold, aFont.Bold
Config.SetSetting ConfigSettingFontItalic, aFont.Italic
Config.SetSetting ConfigSettingFontSize, aFont.Size
Config.SetSetting ConfigSettingFontStrikethrough, aFont.Strikethrough
Config.SetSetting ConfigSettingFontUnderline, aFont.Underline

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sets up logging to write all logged events for all infotypes to a default
' log file.
'
' @remarks
'   The logging level is governed by the global <code>DefaultLogLevel</code> property,
'   but can be overridden if the application's command line specifies the <code>/loglevel</code>
'   switch, which can take any of the following values:
'   <ul>
'       <li>None    or 0</li>
'       <li>Severe  or S</li>
'       <li>Warning or W</li>
'       <li>Info    or I</li>
'       <li>Normal  or N</li>
'       <li>Detail  or D</li>
'       <li>Medium  or M</li>
'       <li>High    or H</li>
'       <li>All     or A</li>
'   </ul>
'
'   If the application's command line includes the <code>/log</code> switch, the
'   the switch's value specifies the log file path and filename.
'
'   Otherwise, the log file is called <code>&lt;appName&gt;.log</code> (where &lt;appName&gt;
'   is the value of the global <code>ApplicationName</code> property). The log file is
'   created in the application settings folder.
' @return
'   The path and filename of the log file.
' @param pCommandLine
'   The command line arguments to be parsed. For a Visual Basic 6 program,
'   this value may be obtained using the <code>Command</code> function.
' @param pOverwriteExisting
'   If <code>True</code>, an existing log file of the same name is overwritten. Otherwise
'   new log entries are appended to an existing log file of the same name (if none exists,
'   a new log file is created).
' @param pCreateBackup
'   If <code>True</code>, a backup copy of an existing log file of the same name is created.
'@/
Public Function SetupDefaultLogging( _
                ByVal pCommandLine As String, _
                Optional ByVal pOverwriteExisting As Boolean = True, _
                Optional ByVal pCreateBackup As Boolean = False) As String
Const ProcName As String = "SetupDefaultLogging"
On Error GoTo Err

If gGetCommandLineParser(pCommandLine).Switch(SwitchLogLevel) Then
    DefaultLogLevel = LogLevelFromString(gGetCommandLineParser(pCommandLine).SwitchValue(SwitchLogLevel))
ElseIf DefaultLogLevel <> LogLevels.LogLevelNormal Then
Else
    DefaultLogLevel = LogLevels.LogLevelNormal
End If

gDefaultLogListener = CreateFileLogListener(DefaultLogFileName(pCommandLine), _
                                        CreateBasicLogFormatter, _
                                        pOverwriteExisting, _
                                        pCreateBackup)

' ensure log entries of all infotypes get written to the log file
gLogManager.GetLogger("").AddLogListener gDefaultLogListener

gLogger.Log App.EXEName & " V" & App.Major & "." & App.Minor & "." & App.Revision, ProcName, ModuleName
gLogger.Log "Log level: " & gLogLevelToString(DefaultLogLevel), ProcName, ModuleName
gLogger.Log "Log file:  " & DefaultLogFileName(pCommandLine), ProcName, ModuleName
SetupDefaultLogging = DefaultLogFileName(pCommandLine)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Sorts an array of <code>Double</code>s into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
' @param Data
'   The array to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
' @param startIndex
'   The index of the first element of the array to be sorted. If not specified,
'   sorting starts at the first element of the <code>data</code> argument.
' @param endIndex
'   The index of the last element of the array to be sorted. If not specified,
'   sorting ends at the last element of the <code>data</code> argument.
' @see
'
'@/
Public Sub SortDoubles( _
                ByRef Data() As Double, _
                Optional ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "SortDoubles"
On Error GoTo Err

If descending Then
    gSortDoublesDesc Data, startIndex, endIndex
Else
    gSortDoublesAsc Data, startIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sorts an array of <code>Long</code>s into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
' @param Data
'   The array to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
' @param startIndex
'   The index of the first element of the array to be sorted. If not specified,
'   sorting starts at the first element of the <code>data</code> argument.
' @param endIndex
'   The index of the last element of the array to be sorted. If not specified,
'   sorting ends at the last element of the <code>data</code> argument.
' @see
'
'@/
Public Sub SortLongs( _
                ByRef Data() As Long, _
                Optional ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "SortLongs"
On Error GoTo Err

If descending Then
gSortLongsDesc Data, startIndex, endIndex
Else
    gSortLongsAsc Data, startIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sorts an array of <code>Object</code>s that implement the <code>Comparable</code>
' interface into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
'
'   Visual Basic does not allow an array of objects declared with a type other than
'   <code>Object</code> to be passed as the <code>data</code> argument. To
'   sort such an array use the <code>SortTypedObjects</code> method.
'
' @param Data
'   The array to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
'   in ascending order.
' @param startIndex
'   The index of the first element of the array to be sorted. If not specified,
'   sorting starts at the first element of the <code>data</code> argument.
' @param endIndex
'   The index of the last element of the array to be sorted. If not specified,
'   sorting ends at the last element of the <code>data</code> argument.
' @see
'
'@/
Public Sub SortObjects( _
                ByRef Data() As Object, _
                Optional ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "SortObjects"
On Error GoTo Err

If descending Then
    gSortObjectsDesc Data, startIndex, endIndex
Else
    gSortObjectsAsc Data, startIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sorts an array of <code>Single</code>s into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
' @param Data
'   The array to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
'   in ascending order.
' @param startIndex
'   The index of the first element of the array to be sorted. If not specified,
'   sorting starts at the first element of the <code>data</code> argument.
' @param endIndex
'   The index of the last element of the array to be sorted. If not specified,
'   sorting ends at the last element of the <code>data</code> argument.
' @see
'
'@/
Public Sub SortSingles( _
                ByRef Data() As Single, _
                Optional ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "SortSingles"
On Error GoTo Err

If descending Then
    gSortSinglesDesc Data, startIndex, endIndex
Else
    gSortSinglesAsc Data, startIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sorts an array of <code>String</code>s into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
' @param Data
'   The array to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
'   in ascending order.
' @param startIndex
'   The index of the first element of the array to be sorted. If not specified,
'   sorting starts at the first element of the <code>data</code> argument.
' @param endIndex
'   The index of the last element of the array to be sorted. If not specified,
'   sorting ends at the last element of the <code>data</code> argument.
' @param caseSensitive
'   Indicates whether a case-sensitive comparison is used during the sorting.
' @see
'
'@/
Public Sub SortStrings( _
                ByRef Data() As String, _
                Optional ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long, _
                Optional ByVal caseSensitive As Boolean)
Const ProcName As String = "SortStrings"
On Error GoTo Err

Dim compareMode As VbCompareMethod

If caseSensitive Then
    compareMode = vbBinaryCompare
Else
    compareMode = vbTextCompare
End If

If descending Then
    gSortStringsDesc Data, compareMode, startIndex, endIndex
Else
    gSortStringsAsc Data, compareMode, startIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sorts an array of objects that implement the <code>Comparable</code>
' interface into ascending order.
'
' @remarks
'   A sub-array of the specified array can be sorted by using the
'   <code>startIndex</code> and/or <code>endIndex</code> arguments.
'
'   This method is used to sort arrays of objects that have not been declared using
'   the <code>Object</code> type.
'
'   Calling this method involves the use of the Visual Basic <code>VarPtr</code>
'   function. For example:
'
'   <pre>
'   <tt>
'   Dim ar(0) As MyClass<br />
'   .<br />
'   .<br />
'   .<br />
'   SortObjects ar      ' doesn't work - causes a type mismatch error<br />
'   SortTypedObjects VarPtr(ar(0)), UBound(ar)+1  ' this works<br />
'   SortTypedObjects VarPtr(ar(5)), 12  ' this sorts 12 elements of the<br />
'                                       ' array starting at the sixth
'   </tt>
'   </pre>
' @param objArrayPointer
'   The address of the first element of the array to be sorted (note that it is
'   necessary to use the Visual Basic <code>VarPtr</code> function to set this
'   argument).
' @param number
'   The number of elements to be sorted.
' @param descending
'   If <code>True</code>, the array is sorted in descending order, otherwise
'   in ascending order.
'@/
Public Sub SortTypedObjects( _
                ByVal objArrayPointer As Long, _
                ByVal number As Long, _
                Optional ByVal descending As Boolean)
Const ProcName As String = "SortTypedObjects"
On Error GoTo Err

gSortTypedObjects objArrayPointer, number, descending, 0, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartSubclassing(ByVal pControl As ISubclassable)
Const ProcName As String = "StartSubclassing"
On Error GoTo Err

gStartSubclassing pControl

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopSubclassing(ByVal hWnd As Long)
Const ProcName As String = "StopSubclassing"
On Error GoTo Err

gStopSubclassing hWnd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Starts a task.
'
' @remarks
'   Note that the first call to the task's <code>Run</code> method does not occur until
'   the event handler that called <code>StartTask</code> has exited.
' @return
'   A <code>TaskController</code> object via which the caller can monitor Progress
'   and completion of the task, or request its cancellation.
' @param Target
'   The task to be started. This is a reference to an object that implements
'   the <code>ITask</code> interface.
' @param Priority
'   Specifies the amount of processor time to be allocated to the task relative
'   to other tasks that may be running. Tasks with priority <code>PriorityNormal</code>
'   get twice as much processor time as tasks with priority <code>PriorityLow</code>,
'   and tasks with priority <code>PriorityHigh</code> get twice as much processor time
'   as tasks with priority <code>PriorityNormal</code>,
' @param Name
'   The name to be given to this task. If not supplied, a GUID is allocated as
'   the task's name.
' @param Cookie
'   A value that is returned unchanged in the <code>TaskProgressEventData</code> and
'   <code>TaskCompletionEventData</code>. The application may use it, for example, as
'   an index into a table of outstanding tasks to avoiding having to search the table
'   to locate an entry when processing these events.
' @param Data
'   Any data that may be needed by the task.
' @see
'
'@/
Public Function StartTask( _
                ByVal Target As ITask, _
                ByVal Priority As TaskPriorities, _
                Optional ByVal Name As String, _
                Optional ByVal Cookie As Variant, _
                Optional ByVal Data As Variant) As TaskController
Const ProcName As String = "StartTask"
On Error GoTo Err

Select Case Priority
Case PriorityNormal
Case PriorityLow
Case PriorityHigh
Case Else
    gAssertArgument False, "Priority invalid"
End Select

Set StartTask = GTasks.gTaskManager.StartTask(Target, Priority, Name, Cookie, Data)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Compares two string arrays for equality.
'
' @remarks
'   Two be considered equal, the arrays must contain the same strings in the
'   same order.
' @return
'   <code>True</code>description if the arrays are equal, <code>False</code> if not.
' @param ar1
'   The first array.
' @param ar2
'   The second array.
' @param caseSensitive
'   If <code>True</code>, a case-sensitive comparison is used when checking equality.
'@/
Public Function StringArraysEqual( _
                ar1() As String, _
                ar2() As String, _
                Optional ByVal caseSensitive As Boolean) As Boolean
Const ProcName As String = "StringArraysEqual"
On Error GoTo Err

Dim i As Long
Dim compareMode As VbCompareMethod

On Error Resume Next    ' in case one array has no members
If UBound(ar1) <> UBound(ar2) Then Exit Function
On Error GoTo Err

If caseSensitive Then
    compareMode = vbBinaryCompare
Else
    compareMode = vbTextCompare
End If

For i = 0 To UBound(ar1)
    If StrComp(ar1(i), ar2(i), compareMode) <> 0 Then Exit Function
Next
StringArraysEqual = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a string to a hexadecimal string.
'
' @return
'   the hexadecimal string corresponding to the supplied
'   string
' @param instring
'   the string to be converted
'@/
Public Function StringToHexString( _
                ByVal inString As String) As String
Const ProcName As String = "StringToHexString"
On Error GoTo Err

StringToHexString = gStringToHexString(inString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' This method may be called prior to ending a program that uses this component. Failure
' to do so may, but usually does not, result in a general protection fault, especially
' when a program is being Run in the Visual Basic 6 development environment.
'@/
Public Sub TerminateTWUtilities()
On Error Resume Next
Globals.gTerminate
End Sub

Public Function TimePeriodFromString( _
                ByVal Value As String) As TimePeriod
Const ProcName As String = "TimePeriodFromString"
On Error GoTo Err

Dim i As Long
Dim Length As Long

Value = Trim$(Value)

i = 1
Do While IsNumeric(Mid$(Value, i, 1))
    i = i + 1
Loop

If i = 1 Then
    Length = 1
Else
    Length = CLng(Left$(Value, i - 1))
End If
    
Set TimePeriodFromString = gGetTimePeriod(Length, _
                                        TimePeriodUnitsFromString(Trim$(Right$(Value, Len(Value) - i + 1))))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits
Const ProcName As String = "TimePeriodUnitsFromString"
On Error GoTo Err

TimePeriodUnitsFromString = gTimePeriodUnitsFromString(timeUnits)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String
Const ProcName As String = "TimePeriodUnitsToString"
On Error GoTo Err

TimePeriodUnitsToString = gTimePeriodUnitsToString(timeUnits)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TimePeriodUnitsToShortString( _
                timeUnits As TimePeriodUnits) As String
Const ProcName As String = "TimePeriodUnitsToShortString"
On Error GoTo Err

TimePeriodUnitsToShortString = gTimePeriodUnitsToShortString(timeUnits)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Writes a trace entry, using the null trace type, to the logging system recording
' entry to the specified procedure.
'
' @remarks
'   Using this method avoids the need for the application to maintain a reference to
'   the null tracer object.
' @param pProcedureName
'   The name of the procedure in which the method is called.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pInfo
'   Any additional information to be logged, such as argument values.
'@/
Public Sub TraceProcedureEntry( _
                ByVal pProcedureName As String, _
                Optional ByVal pProjectName As String, _
                Optional ByVal pModuleName As String, _
                Optional ByVal pInfo As String)
gNullTracer.EnterProcedure pProcedureName, pProjectName, pModuleName, pInfo
End Sub

''
' Writes a trace entry, using the null trace type, to the logging system recording
' exit from the specified procedure.
'
' @remarks
'   Using this method avoids the need for the application to maintain a reference to
'   the null tracer object.
' @param pProcedureName
'   The name of the procedure in which the method is called.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pInfo
'   Any additional information to be logged, such as argument values.
'@/
Public Sub TraceProcedureExit( _
                ByVal pProcedureName As String, _
                Optional ByVal pProjectName As String, _
                Optional ByVal pModuleName As String, _
                Optional ByVal pInfo As String)
gNullTracer.ExitProcedure pProcedureName, pProjectName, pModuleName, pInfo
End Sub

''
' Returns a string representation of a Variant value.
'
' @remarks
'   This function takes account of the variant type, including whether it
'   is an array.
'
'   For variants whose type is permitted for use with the VB <code>Cstr</code> function,
'   the returned string is just the output from <code>Cstr</code>.
'
'   For object variants:
'   <ul>
'   <li>if the object implements the <code>Stringable</code> interface,
'   the returned string includes the output from the <code>ToString</code> method.</li>
'   <li>if the object implements the <code>JSONable</code> interface,
'   the returned string includes the output from the <code>ToJSON</code> method.</li>
'   </ul>
'
' @return
'   A string representation of the supplied Variant value.
' @param Value
'   The value whose string representation is required.
'@/
Public Function VariantToString( _
                ByVal Value As Variant) As String
Const ProcName As String = "VariantToString"
On Error GoTo Err

VariantToString = gVariantToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Pauses execution for the specified period, while allowing events to be
' processed as normal.
'
' @param waitTimeMillisecs
'   Specifies the duration for which execution is to be paused, in milliseconds.
'@/
Public Sub Wait( _
                ByVal waitTimeMillisecs As Long)
Const ProcName As String = "Wait"
On Error GoTo Err

doWait CCur(-waitTimeMillisecs)
 
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Pauses execution for the specified period, while allowing events to be
' processed as normal.
'
' @param waitUntil
'   Specifies the time at which execution is to resume.
'@/
Public Sub WaitUntilTime( _
                ByVal waitUntil As Date)
Const ProcName As String = "WaitUntilTime"
On Error GoTo Err

gAssertArgument waitUntil >= Now, "Specified time is earlier than now"

doWait gVbDateToFileTime(gLocalToUtc(waitUntil))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function WeekStartDate( _
                ByVal pDate As Date) As Date
Const ProcName As String = "WeekStartDate"
On Error GoTo Err

WeekStartDate = gCalcWeekStartDate(pDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function WeekStartDateFromWeekNumber( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "WeekStartDateFromWeekNumber"
On Error GoTo Err

WeekStartDateFromWeekNumber = gCalcWeekStartDateFromWeekNumber(weekNumber, baseDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function WorkingDayDate( _
                ByVal dayNumber As Long, _
                ByVal baseDate As Date) As Date
Const ProcName As String = "WorkingDayDate"
On Error GoTo Err

WorkingDayDate = gCalcWorkingDayDate(dayNumber, baseDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function WorkingDayNumber( _
                ByVal pDate As Date) As Long
Const ProcName As String = "WorkingDayNumber"
On Error GoTo Err

WorkingDayNumber = gCalcWorkingDayNumber(pDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doWait( _
                ByVal dueTime As Currency)
Const ProcName As String = "doWait"
On Error GoTo Err

Dim hTimer As Long
Dim ret As Long

hTimer = CreateWaitableTimer(0, True, "")

ret = SetWaitableTimer(hTimer, dueTime, 0, 0, 0, False)

Do
    ret = MsgWaitForMultipleObjects(1, _
                                    hTimer, _
                                    False, _
                                    INFINITE, _
                                    QS_ALLEVENTS)
    gAssert ret <> WAIT_FAILED, "MsgWaitForMultipleObjects failed with Error code: " & Err.LastDllError, ErrorCodes.ErrRuntimeException

    DoEvents
Loop Until ret = WAIT_OBJECT_0

CloseHandle hTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub










