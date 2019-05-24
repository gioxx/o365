' Name: Robust Office Inventory Scan - Version 1.9.1
' Author: Microsoft Customer Support Services
' Copyright (c) Microsoft Corporation. All rights reserved.
' Script to create an inventory scan of installed Office applications
' Supported Office Families: 2000, 2002, 2003, 2007
'                            2010, 2013, 2016, O365

Option Explicit
On Error Resume Next
Const SCRIPTBUILD = "1.9.1"
Dim sPathOutputFolder : sPathOutputFolder = ""
Dim fQuiet : fQuiet = False
Dim fLogFeatures : fLogFeatures = False
Dim fLogFull : fLogFull = False
dim fBasicMode : fBasicMode = False
Dim fLogChainedDetails : fLogChainedDetails = False
Dim fLogVerbose : fLogVerbose = False
Dim fListNonOfficeProducts : fListNonOfficeProducts = False
Dim fFileInventory : fFileInventory = False
Dim fFeatureTree : fFeatureTree = False
Dim fDisallowCScript : fDisallowCScript = False

'=======================================================================================================
'[INI] Section for script behavior customizations

'Directory for Log output.
'Example: "\\<server>\<share>\"
'Default: sPathOutputFolder = vbNullString -> %temp% directory is used
sPathOutputFolder = ""

'Quiet switch.
'Default: False -> Open inventory log when done
fQuiet = False


'Basic Mode
'Generates a basic list of installed Office products with licensing information only
'Disables all other extended analysis options
'Default: False -> Allow extended analysis options
fBasicMode = False

'Log full (verbose) details. This enables all possible scans for Office products.
'Default: False -> Only list standard details
fLogFull = False

'Enables additional logging details
'like the "FeatureTree" display and
'the additional individual listing of the chained Office SKU's
'Default: False -> Do not log additional details
fLogVerbose = False

'Starting with Office 2007 a SKU can contain several .msi packages
'The default does not list the details for each chained package
'This option allows to show the full details for each chained package
'Default: Comprehensive view - do not show details for chained packages
fLogChainedDetails = False

'The script filters for products of the Microsoft Office family.
'Set this option to 'True' to get a list of all Windows Installer products in the inventory log
'Default: False -> Don't list other products in the log
fListNonOfficeProducts = False

'File level inventory of installed Office products
'Depending on the number of installed products this can be an extremely time consuming task!
'Default: False -> Don't create a file level inventory
fFileInventory = False

'Detect all features of a product and include a feature tree in the log
'Default: False -> Don't include the feature detection
fFeatureTree = False

'DO NOT CUSTOMIZE BELOW THIS LINE!
'=======================================================================================================

'Measure total scan runtime
Dim tStart, tEnd
tStart = Time()
'Call the command line parser
ParseCmdLine
'Definition of non customizable settings
'Strings 
Dim sComputerName, sTemp, sCurUserSid, sDebugErr, sError, sErrBpa, sProductCodes_C2R
Dim sStack, sCacheLog, sLogFile, sInstalledProducts, sSystemType, sLogFormat
Dim sPackageGuid, sUserName, sDomain
'Arrays
Dim arrAllProducts(), arrMProducts(), arrUUProducts(), arrUMProducts(), arrMaster(), arrArpProducts()
Dim arrVirtProducts(), arrVirt2Products(), arrVirt3Products(), arrPatch(), arrAipPatch, arrMspFiles
Dim arrLog(4), arrLogFormat(), arrUUSids(), arrUMSids(), arrMVProducts(), arrIS(), arrFeature()
Dim arrProdVer09(), arrProdVer10(), arrProdVer11(), arrProdVer12(), arrProdVer14(), arrProdVer15()
Dim arrProdVer16(), arrFiles(), arrLicenese()
'Booleans
Dim fIsAdmin, fIsElevated, fIsCriticalError, fGuidCaseWarningOnly, f64, fPatchesOk, fPatchesExOk
Dim fCScript, bOsppInit, fZipError, fInitArrProdVer, fv2SxS 
'Integers
Dim iWiVersionMajor, iVersionNt, iPCount, iPatchesExError, iVMVirtOverride
'Dictionaries
Dim dicFolders, dicAssembly, dicMspIndex, dicProducts, dicArp, dicMissingChild
Dim dicPatchLevel, dicScenario, dicKeyComponents
Dim dicPolHKCU, dicPolHKLM
Dim dicProductCodeC2R, dicActiveC2RVersions, dicMapArpToConfigID
Dim dicKeyComponentsV2, dicScenarioV2, dicC2RPropV2, dicVirt2Cultures
Dim dicKeyComponentsV3, dicScenarioV3, dicC2RPropV3, dicVirt3Cultures
'Other
Dim oMsi, oShell, oFso, oReg, oWsh, oWMILocal
Dim TextStream, ShellApp, AppFolder, Ospp, Spp

'Identifier for product family
Const OFFICE_ALL                      = "78E1-11D2-B60F-006097C998E7}.0001-11D2-92F2-00104BC947F0}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.7000-11D3-8CFE-0150048383C9}.BE5F-4ED1-A0F7-759D40C7622E}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}.CFDA-404E-8992-6AF153ED1719}.{9AC08E99-230B-47e8-9721-4577B7F124EA}"
'Office 2000 -> KB230848; Office XP -> KB302663; Office 2003 -> KB832672
Const OFFICE_2000                     = "78E1-11D2-B60F-006097C998E7}"
Const ORK_2000                        = "0001-11D2-92F2-00104BC947F0}"
Const PRJ_2000                        = "BDCA-11D1-B7AE-00C04FB92F3D}"
Const VIS_2002                        = "6D54-11D4-BEE3-00C04F990354}"
Const OFFICE_2002                     = "6000-11D3-8CFE-0050048383C9}"
Const OFFICE_2003                     = "6000-11D3-8CFE-0150048383C9}"
Const WSS_2                           = "7000-11D3-8CFE-0150048383C9}"
Const SPS_2003                        = "BE5F-4ED1-A0F7-759D40C7622E}"
Const PPS_2007                        = "CFDA-404E-8992-6AF153ED1719}" 'Project Portfolio Server 2007
Const POWERPIVOT_2010                 = "{72F8ECCE-DAB0-4C23-A471-625FEDABE323}, {A37E1318-29CA-4A9F-9CCA-D9BFDD61D17B}" 'UpgradeCode!
Const O15_C2R                         = "{9AC08E99-230B-47e8-9721-4577B7F124EA}"
Const OFFICEID                        = "000-0000000FF1CE}" 'cover O12, O14 with 32 & 64 bit
Const OREGREFC2R15                    = "Microsoft Office 15"

'SPP AppId
Const OFFICE14APPID                   = "59a52881-a989-479d-af46-f275c6370663"
Const OFFICEAPPID                     = "0ff1ce15-a989-479d-af46-f275c6370663"

Const PRODLEN                         = 13
Const FOR_READING                     = 1
Const FOR_WRITING                     = 2
Const FOR_APPENDING                   = 8
Const TRISTATE_USEDEFAULT             = -2 'Opens the file using the system default. 
Const TRISTATE_TRUE                   = -1 'Opens the file as Unicode. 
Const TRISTATE_FALSE                  = 0  'Opens the file as ASCII. 

Const USERSID_EVERYONE                = "s-1-1-0"
Const MACHINESID                      = ""
Const PRODUCTCODE_EMPTY               = ""
Const MSIOPENDATABASEMODE_READONLY    = 0
Const MSIOPENDATABASEMODE_PATCHFILE   = 32
Const MSICOLUMNINFONAMES              = 0
Const MSICOLUMNINFOTYPES              = 1
'Summary Information fields
Const PID_TITLE                       = 2 'Type of installer package. E.g. "Installation Database" or "Transform" or "Patch"
Const PID_SUBJECT                     = 3 'Displayname
Const PID_TEMPLATE                    = 7 'compatible platform and language versions for .msi / PatchTargets for .msp
Const PID_REVNUMBER                   = 9 'PackageCode
Const PID_WORDCOUNT                   = 15'InstallSource type 
Const MSIPATCHSTATE_UNKNOWN           = -1 'Patch is in an unknown state to this product instance. 
Const MSIPATCHSTATE_APPLIED           = 1 'Patch is applied to this product instance. 
Const MSIPATCHSTATE_SUPERSEDED        = 2 'Patch is applied to this product instance but is superseded.  
Const MSIPATCHSTATE_OBSOLETED         = 4 'Patch is applied in this product instance but obsolete.  
Const MSIPATCHSTATE_REGISTERED        = 8 'The enumeration includes patches that are registered but not yet applied.
Const MSIPATCHSTATE_ALL               = 15
Const MSIINSTALLCONTEXT_USERMANAGED   = 1
Const MSIINSTALLCONTEXT_USERUNMANAGED = 2
Const MSIINSTALLCONTEXT_MACHINE       = 4
Const MSIINSTALLCONTEXT_ALL           = 7
Const MSIINSTALLCONTEXT_C2RV2         = 8 'C2r V2 virtualized context
Const MSIINSTALLCONTEXT_C2RV3         = 15 'C2r V3 virtualized context
Const MSIINSTALLMODE_DEFAULT          = 0    'Provide the component and perform any installation necessary to provide the component. 
Const MSIINSTALLMODE_EXISTING         = -1   'Provide the component only if the feature exists. This option will verify that the assembly exists.
Const MSIINSTALLMODE_NODETECTION      = -2   'Provide the component only if the feature exists. This option does not verify that the assembly exists.
Const MSIINSTALLMODE_NOSOURCERESOLUTION = -3 'Provides the assembly only if the assembly is installed local.
 Const MSIPROVIDEASSEMBLY_NET         = 0    'A .NET assembly.
Const MSIPROVIDEASSMBLY_WIN32         = 1    'A Win32 side-by-side assembly.
Const MSITRANSFORMERROR_ALL           = 319
'Installstates for products, features, components
Const INSTALLSTATE_NOTUSED            = -7 ' component disabled
Const INSTALLSTATE_BADCONFIG          = -6 ' configuration data corrupt
Const INSTALLSTATE_INCOMPLETE         = -5 ' installation suspended or in progress
Const INSTALLSTATE_SOURCEABSENT       = -4 ' run from source, source is unavailable
Const INSTALLSTATE_MOREDATA           = -3 ' return buffer overflow
Const INSTALLSTATE_INVALIDARG         = -2 ' invalid function argument. The product/feature is neither advertised or installed. 
Const INSTALLSTATE_UNKNOWN            = -1 ' unrecognized product or feature
Const INSTALLSTATE_BROKEN             =  0 ' broken
Const INSTALLSTATE_ADVERTISED         =  1 ' The product/feature is advertised but not installed. 
Const INSTALLSTATE_REMOVED            =  1 ' The component is being removed (action state, not settable)
Const INSTALLSTATE_ABSENT             =  2 ' The product/feature is not installed. 
Const INSTALLSTATE_LOCAL              =  3 ' The product/feature/component is installed. 
Const INSTALLSTATE_SOURCE             =  4 ' The product or feature is installed to run from source, CD, or network. 
Const INSTALLSTATE_DEFAULT            =  5 ' The product or feature will be installed to use the default location: local or source.
Const INSTALLSTATE_VIRTUALIZED        =  8 ' The product is virtualized (C2R).
Const VERSIONCOMPARE_LOWER            = -1 ' Left hand file version is lower than right hand 
Const VERSIONCOMPARE_MATCH            =  0 ' File versions are identical
Const VERSIONCOMPARE_HIGHER           =  1 ' Left hand file versin is higher than right hand
Const VERSIONCOMPARE_INVALID          =  2 ' Cannot compare. Invalid compare attempt.

Const COPY_OVERWRITE                  = &H10&
Const COPY_SUPPRESSERROR              = &H400& 

Const HKEY_CLASSES_ROOT               = &H80000000
Const HKEY_CURRENT_USER               = &H80000001
Const HKEY_LOCAL_MACHINE              = &H80000002
Const HKEY_USERS                      = &H80000003
Const HKCR                            = &H80000000
Const HKCU                            = &H80000001
Const HKLM                            = &H80000002
Const HKU                             = &H80000003
Const KEY_QUERY_VALUE                 = &H0001
Const KEY_SET_VALUE                   = &H0002
Const KEY_CREATE_SUB_KEY              = &H0004
Const DELETE                          = &H00010000
Const REG_SZ                          = 1
Const REG_EXPAND_SZ                   = 2
Const REG_BINARY                      = 3
Const REG_DWORD                       = 4
Const REG_MULTI_SZ                    = 7
Const REG_QWORD                       = 11
Const REG_GLOBALCONFIG                = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\"
Const REG_CONTEXTMACHINE              = "Installer\"
Const REG_CONTEXTUSER                 = "Software\Microsoft\Installer\"
Const REG_CONTEXTUSERMANAGED          = "Software\Microsoft\Windows\CurrentVersion\Installer\Managed\"
Const REG_ARP                         = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const REG_OFFICE                      = "SOFTWARE\Microsoft\Office\"
Const REG_C2RVIRT_HKLM                = "\ClickToRun\REGISTRY\MACHINE\"

Const STR_NOTCONFIGURED             = "<Not Configured>"
Const STR_NOTCONFIGUREDXML          = "NotConfigured"
Const STR_PACKAGEGUID               = "PackageGUID"
Const STR_REGPACKAGEGUID            = "RegPackageGUID"
Const STR_BUILDNUMBER               = "BuildNumber"
Const STR_VERSIONTOREPORT           = "BuildVersionToReport"
Const STR_VERSION                   = "Version"
Const STR_PLATFORM                  = "Platform"
Const STR_OFFICEMGMTCOM             = "OfficeMgmtCom"
Const STR_CDNBASEURL                = "CDNBaseUrl"
Const STR_LASTUSEDBASEURL           = "Last used InstallSource"
Const STR_UPDATELOCATION            = "Custom UpdateLocation"
Const STR_USEDUPDATELOCATION        = "Winning UpdateLocation"
Const STR_UPDATESENABLED            = "UpdatesEnabled"
Const STR_UPDATECHANNEL             = "UpdateChannel"
Const STR_UPDATETOVERSION           = "UpdateToVersion"
Const STR_UPDATETHROTTLE            = "UpdatesThrottleValue"
Const STR_POLUPDATESENABLED         = "Policy UpdatesEnabled"
Const STR_POLUPGRADEENABLED         = "Policy EnableAutomaticUpgrade"
Const STR_POLUPDATECHANNEL          = "Policy UpdateChannel"
Const STR_POLUPDATELOCATION         = "Policy UpdateLocation"
Const STR_POLUPDATETOVERSION        = "Policy UpdateToVersion"
Const STR_POLUPDATEDEADLINE         = "Policy UpdateDeadline"
Const STR_POLUPDATENOTIFICATIONS    = "Policy UpdateHideNotifications"
Const STR_POLHIDEUPDATECFGOPT       = "Policy HideUpdateConfigOptions"
Const STR_SCA                       = "Shared Computer Licensing"
Const STR_POLSCACACHEOVERRIDE       = "Policy SCLCacheOverride"
Const STR_POLSCACACHEOVERRIDEDIR    = "Policy SCLCacheOverrideDirectory"
Const STR_POLOFFICEMGMTCOM          = "Policy OfficeMgmtCom"

Const GUID_UNCOMPRESSED               = 0
Const GUID_COMPRESSED                 = 1
Const GUID_SQUISHED                   = 2
Const LOGPOS_COMPUTER                 = 0 '    ArrLogPosition 0: "Computer"
Const LOGPOS_REVITEM                  = 1 '    ArrLogPosition 1: "Review Items"
Const LOGPOS_PRODUCT                  = 2 '    ArrLogPosition 2: "Products Inventory"
Const LOGPOS_RAW                      = 3 '    ArrLogPosition 3: "Raw Data"
Const LOGHEADING_NONE                 = 0 '    Not a heading
Const LOGHEADING_H1                   = 1 '    Heading 1 '='
Const LOGHEADING_H2                   = 2 '    Heading 2 '-'
Const LOGHEADING_H3                   = 3 '    Heading 3 ' '
Const TEXTINDENT                      = "                        "
Const CATEGORY                        = 1
Const TAG                             = 2

'Global_Access_Core - msaccess.exe
Const CID_ACC16_64 = "{27C919A6-3FA5-47F9-A3EC-BC7FF2AAD452}"
Const CID_ACC16_32 = "{E34AA7C4-8845-4BD7-BAC6-26554B60823B}"
Const CID_ACC15_64 = "{3CE2B4B3-DA38-4113-8DB2-965847CDE94F}"
Const CID_ACC15_32 = "{A3E12EF0-7C3B-4493-99A3-F92FCD0AA512}"
Const CID_ACC14_64 = "{02F5CBEC-E7B5-4FC1-BD72-6043152BD1D4}"
Const CID_ACC14_32 = "{AE393348-E564-4894-B8C5-EBBC5E72EFC6}"
Const CID_ACC12    = "{0638C49D-BB8B-4CD1-B191-054E8F325736}"
Const CID_ACC11    = "{F2D782F8-6B14-4FA4-8FBA-565CDDB9B2A8}"
'Global_Excel_Core - excel.exe
Const CID_XL16_64 = "{C4ACE6DB-AA99-401F-8BE6-8784BD09F003}"
Const CID_XL16_32 = "{C845E028-E091-442E-8202-21F596C559A0}"
Const CID_XL15_64 = "{58A9998B-6103-436F-85A1-52720802CA0A}"
Const CID_XL15_32 = "{107E1A9A-03AE-4F2B-ACF7-0CC519E60E7B}"
Const CID_XL14_64 = "{8B1BF0B4-A1CA-4656-AA46-D11C50BC55A4}"
Const CID_XL14_32 = "{538F6C89-2AD5-4006-8154-C6670774E980}"
Const CID_XL12    = "{0638C49D-BB8B-4CD1-B191-052E8F325736}"
Const CID_XL11    = "{A2B280D4-20FB-4720-99F7-40C09FBCE10A}"
'WAC_CoreSPD - spdesign.exe (frontpage.exe)
Const CID_SPD16_64 = "{2FB768AF-8F57-424A-BBDA-81611CFF3ED2}"
Const CID_SPD16_32 = "{C3F352B2-A43B-4948-AE54-12E265647697}"
Const CID_SPD15_64 = "{25B4430E-E7D6-406F-8468-D9B65BC240F3}"
Const CID_SPD15_32 = "{0F0A451D-CB3C-44BE-B8A4-E72C2B89C4A2}"
Const CID_SPD14_64 = "{6E4D3AA2-2AD9-4DD2-8C2D-8C55B656A5C9}"
Const CID_SPD14_32 = "{E5344AC3-915E-4655-AF0D-98BC878805DC}"
Const CID_SPD12    = "{0638C49D-BB8B-4CD1-B191-056E8F325736}"
Const CID_SPD11    = "{81E9830C-5A6B-436A-BEC9-4FB759282DE3}" ' FrontPage
'Groove_Core - groove.exe
Const CID_GRV16_64 = "{EEE31981-E2D9-45AE-B134-FD9276C19588}"
Const CID_GRV16_32 = "{6C26357C-A2D8-4C68-8BC6-A8091BECDA02}"
Const CID_GRV15_64 = "{AD8AD7F2-98CB-4257-BE7A-05CBCA1354B4}"
Const CID_GRV15_32 = "{87E86C36-1368-4841-9152-766F31BC46E8}"
Const CID_GRV14_64 = "{61CD70FF-C6B7-4F6A-8491-5B8B9B0040F8}"
Const CID_GRV14_32 = "{EFE67578-E52B-410E-9178-9911443DBF5A}"
Const CID_GRV12    = "{0A048D77-2DE9-4672-ACF7-12429662397D}"
'Lync_Corelync - lync.exe
Const CID_LYN16_64 = "{3CFF5AB2-9B16-4A31-BC3F-FAD761D92780}"
Const CID_LYN16_32 = "{E1AFBCD9-12F0-4FC0-9177-BFD3148AEC74}"
Const CID_LYN15_64 = "{D5B16A67-9FA6-4B77-AE2A-3B1F49CE9D3B}"
Const CID_LYN15_32 = "{F8D36F1C-6196-4FFA-94AA-736644D458E3}"
'Global_OneNote_Core - onenote.exe
Const CID_ONE16_64 = "{8265A5EF-46C7-4D46-812C-076F2A28F7CB}"
Const CID_ONE16_32 = "{2A8FA8D7-B728-4792-AC02-463FD7A423BD}"
Const CID_ONE15_64 = "{74F233A9-A17A-477C-905F-853F5FCDAD40}"
Const CID_ONE15_32 = "{DACE5A15-C57C-44DE-9AFF-89B4412485AF}"
Const CID_ONE14_64 = "{9542A6E5-2FAF-4191-B525-6ED00F2D0127}"
Const CID_ONE14_32 = "{62F8C897-D359-4D8F-9659-CF1E9E3E6B74}"
Const CID_ONE12    = "{0638C49D-BB8B-4CD1-B191-057E8F325736}"
Const CID_ONE11    = "{D2C0E18B-C463-4E90-92AC-CA94EBEC26CE}"
'Global_Office_Core - mso.dll
Const CID_MSO16_64 = "{625F5772-C1B3-497E-8ABE-7254EDB00506}"
Const CID_MSO16_32 = "{68477CB0-662A-48FB-AF2E-9573C92869F7}"
Const CID_MSO15_64 = "{D01398A1-F26F-4545-A441-567F097A57D7}"
Const CID_MSO15_32 = "{9CC2CF5E-9A2E-41AC-AF95-432890A9659A}"
Const CID_MSO14_64 = "{E6AC97ED-6651-4C00-A8FE-790DB0485859}"
Const CID_MSO14_32 = "{398E906A-826B-48DD-9791-549C649CACE5}"
Const CID_MSO12    = "{0638C49D-BB8B-4CD1-B191-050E8F325736}"
Const CID_MSO11    = "{A2B280D4-20FB-4720-99F7-10C09FBCE10A}"
'Global_Outlook_Core - outlook.exe
Const CID_OL16_64 = "{7C6D92EF-7B45-46E5-8670-819663220E4E}"
Const CID_OL16_32 = "{2C6C511D-4542-4E0C-95D0-05D4406032F2}"
Const CID_OL15_64 = "{3A5F96E7-F51D-4942-98DB-3CD037FB39E5}"
Const CID_OL15_32 = "{E9E5CFFC-AFFE-4F83-A695-7734FA4775B9}"
Const CID_OL14_64 = "{ECCC8A38-7855-46CA-88FB-3BAA7CD95E56}"
Const CID_OL14_32 = "{CFF13DD8-6EF2-49EB-B265-E3BFC6501C1D}"
Const CID_OL12    = "{0638C49D-BB8B-4CD1-B191-055E8F325736}"
Const CID_OL11    = "{3CE26368-6322-4ABF-B11B-458F5C450D0F}"
'Global_PowerPoint_Core - powerpnt.exe
Const CID_PPT16_64 = "{E0A76492-0FD5-4EC2-8570-AE1BAA61DC88}"
Const CID_PPT16_32 = "{9E73CEA4-29D0-4D16-8FB9-5AB17387C960}"
Const CID_PPT15_64 = "{8C1B8825-A280-4657-A7B8-8172C553A4C4}"
Const CID_PPT15_32 = "{258D5292-6DDA-4B39-B301-58405FA16638}"
Const CID_PPT14_64 = "{EE8D8E0A-D905-401D-9BC3-0D20156D5E30}"
Const CID_PPT14_32 = "{E72E0D20-0D63-438B-BC71-92AB9F9E8B54}"
Const CID_PPT12    = "{0638C49D-BB8B-4CD1-B191-053E8F325736}"
Const CID_PPT11    = "{C86C0B92-63C0-4E35-8605-281275C21F97}"
'Global_Project_ClientCore - winproj.exe
Const CID_PRJ16_64 = "{107BCD9A-F1DC-4004-A444-33706FC10058}"
Const CID_PRJ16_32 = "{0B6EDA1D-4A15-4F88-8B20-EA6528978E4E}"
Const CID_PRJ15_64 = "{760CE47D-9512-40D9-8C6D-CF232851B4BB}"
Const CID_PRJ15_32 = "{5296AE31-2F7D-480C-BFDC-CE0797426395}"
Const CID_PRJ14_64 = "{64A809BD-6EE9-475C-B4E8-95B0D7FF3B97}"
Const CID_PRJ14_32 = "{51894540-193D-40AE-83F9-D3FC5DB24D91}"
Const CID_PRJ12    = "{43C3CF66-AA31-476D-B029-6D274E46F86C}"
Const CID_PRJ11    = "{C33FFB81-6E54-4541-AFF4-D84DC60460F7}"
'Global_Publisher_Core - mspub.exe
Const CID_PUB16_64 = "{7ECBF2AA-14AA-4F89-B9A5-C064274CFA83}"
Const CID_PUB16_32 = "{81DD86EC-5F1C-4DDE-9211-98AF184EAD47}"
Const CID_PUB15_64 = "{22299AFF-DC4C-45A8-9A8F-651FB6467057}"
Const CID_PUB15_32 = "{C9C0167D-3FE0-4078-B47E-83272A4B8B04}"
Const CID_PUB14_64 = "{A716400F-5D5D-45CF-94B4-05B17A98B901}"
Const CID_PUB14_32 = "{CD0D7B29-89E7-49C5-8EE1-5D858EFF2593}"
Const CID_PUB12    = "{CD0D7B29-89E7-49C5-8EE1-5D858EFF2593}"
Const CID_PUB11    = "{0638C49D-BB8B-4CD1-B191-05CE8F325736}"
'Global_XDocs_Core - infopath.exe
Const CID_IP16_64 = "{2774AAC0-1433-46BE-993F-8088018C3B09}"
Const CID_IP15_64 = "{19AF7201-09A2-4C73-AB50-FCEF94CB2BA9}"
Const CID_IP15_32 = "{3741355B-72CF-4CEE-948E-CC9FBDBB8E7A}"
Const CID_IP14_64 = "{28B2FBA8-B95F-47CB-8F8F-0885ACDAC69B}"
Const CID_IP14_32 = "{E3898C62-6EC3-4491-8194-9C88AD716468}"
Const CID_IP12    = "{0638C49D-BB8B-4CD1-B191-058E8F325736}"
Const CID_IP11    = "{1A66B512-C4BE-4347-9F0C-8638F8D1E6E4}"
'Global_Visio_visioexe - visio.exe
Const CID_VIS16_64 = "{2D4540EC-2C88-4C28-AE88-2614B5460648}"
Const CID_VIS16_32 = "{A4C55BC1-B94C-4058-B15C-B9D4AE540AD1}"
Const CID_VIS15_64 = "{7069FF90-1D63-4F85-A2AB-6F0D01C78D83}"
Const CID_VIS15_32 = "{5D502092-1543-4D9B-89FE-7B4364417CC6}"
Const CID_VIS14_64 = "{DB2B19E4-F894-47B1-A6F1-9B391A4AE0A8}"
Const CID_VIS14_32 = "{4371C2B1-3F27-41F5-A849-9987AB91D990}"
Const CID_VIS12    = "{0638C49D-BB8B-4CD1-B191-05DE8F325736}"
Const CID_VIS11    = "{7E5F9F34-8EA7-4EA2-ABFB-CA4E742EFFA1}"
'Global_Word_Core - winword.exe
Const CID_WD16_64 = "{DC5CCACD-A7AC-4FD3-9F70-9454B5DE5161}"
Const CID_WD16_32 = "{30CAC893-3CA4-494C-A5E9-A99141352216}"
Const CID_WD15_64 = "{6FF09BDF-B087-4E23-A9B9-272DBFD64099}"
Const CID_WD15_32 = "{09D07EFC-505F-4D9C-BFD5-ACE3217F6654}"
Const CID_WD14_64 = "{C0AC079D-A84B-4CBD-8DBA-F1BB44146899}"
Const CID_WD14_32 = "{019C826E-445A-4649-A5B0-0BF08FCC4EEE}"
Const CID_WD12    = "{0638C49D-BB8B-4CD1-B191-051E8F325736}"
Const CID_WD11    = "{1EBDE4BC-9A51-4630-B541-2561FA45CCC5}"

'Arrays
Const UBOUND_LOGARRAYS                = 12
Const UBOUND_LOGCOLUMNS               = 34 ' Controlled by array with the most columns
Redim arrLogFormat(UBOUND_LOGARRAYS, UBOUND_LOGCOLUMNS)

Const ARRAY_MASTER                                          = 0 'Master data array id
    Const UBOUND_MASTER                                     = 34
    Const COL_PRODUCTCODE                                   = 0
    Const COL_PRODUCTNAME                                   = 1 ' Msi ProductName
    Const COL_USERSID                                       = 2
    Const COL_CONTEXTSTRING                                 = 3 ' ProductContext
    Const COL_STATESTRING                                   = 4 ' ProductState
    Const COL_CONTEXT                                       = 5 ' ProductContext
    Const COL_STATE                                         = 6 ' ProductState
    Const COL_SYSTEMCOMPONENT                               = 7 ' Arp SystemComponent
    Const COL_ARPPARENTCOUNT                                = 8
    Const COL_ARPPARENTS                                    = 9
    Const COL_ARPPRODUCTNAME                                = 10
    Const COL_PRODUCTVERSION                                = 11
    Const COL_SPLEVEL                                       = 12 ' ServicePack Level
    Const COL_INSTALLDATE                                   = 13
    Const COL_CACHEDMSI                                     = 14
    Const COL_ORIGINALMSI                                   = 15 ' Original .msi name
    Const COL_ORIGIN                                        = 16 ' Build/Origin Property
    Const COL_PRODUCTID                                     = 17 ' ProductID Property
    Const COL_PACKAGECODE                                   = 18
    Const COL_TRANSFORMS                                    = 19
    Const COL_ARCHITECTURE                                  = 20
    Const COL_ERROR                                         = 21
    Const COL_NOTES                                         = 22
    Const COL_METADATASTATE                                 = 23
    Const COL_ISOFFICEPRODUCT                               = 24
    Const COL_PATCHFAMILY                                   = 25
    Const COL_OSPPLICENSE                                   = 26
    Const COL_OSPPLICENSEXML                                = 27
    Const COL_ACIDLIST                                      = 28
    Const COL_LICENSELIST                                   = 29
    Const COL_UPGRADECODE                                   = 30
    Const COL_VIRTUALIZED                                   = 31
    Const COL_INSTALLTYPE                                   = 32
    Const COL_KEYCOMPONENTS                                 = 33
    Const COL_PRIMARYPRODUCTID                              = 34
    arrLogFormat(ARRAY_MASTER, COL_PRODUCTCODE)             = "ProductCode"
    arrLogFormat(ARRAY_MASTER, COL_PRODUCTNAME)             = "Msi ProductName"
    arrLogFormat(ARRAY_MASTER, COL_USERSID)                 = "UserSid"
    arrLogFormat(ARRAY_MASTER, COL_CONTEXTSTRING)           = "ProductContext"
    arrLogFormat(ARRAY_MASTER, COL_STATESTRING)             = "ProductState"
    arrLogFormat(ARRAY_MASTER, COL_CONTEXT)                 = "ProductContext"
    arrLogFormat(ARRAY_MASTER, COL_STATE)                   = "ProductState"
    arrLogFormat(ARRAY_MASTER, COL_SYSTEMCOMPONENT)         = "Arp SystemComponent"
    arrLogFormat(ARRAY_MASTER, COL_ARPPARENTCOUNT)          = "Arp ParentCount"
    arrLogFormat(ARRAY_MASTER, COL_ARPPARENTS)              = "Configuration SKU"
    arrLogFormat(ARRAY_MASTER, COL_ARPPRODUCTNAME)          = "ARP ProductName"
    arrLogFormat(ARRAY_MASTER, COL_PRODUCTVERSION)          = "ProductVersion"
    arrLogFormat(ARRAY_MASTER, COL_SPLEVEL)                 = "ServicePack Level"
    arrLogFormat(ARRAY_MASTER, COL_INSTALLDATE)             = "InstallDate"
    arrLogFormat(ARRAY_MASTER, COL_CACHEDMSI)               = "Cached .msi Package"
    arrLogFormat(ARRAY_MASTER, COL_ORIGINALMSI)             = "Original .msi Name"
    arrLogFormat(ARRAY_MASTER, COL_ORIGIN)                  = "Build/Origin"
    arrLogFormat(ARRAY_MASTER, COL_PRODUCTID)               = "ProductID (MSI)"
    arrLogFormat(ARRAY_MASTER, COL_PACKAGECODE)             = "Package Code"
    arrLogFormat(ARRAY_MASTER, COL_TRANSFORMS)              = "Transforms"
    arrLogFormat(ARRAY_MASTER, COL_ARCHITECTURE)            = "Architecture"
    arrLogFormat(ARRAY_MASTER, COL_ERROR)                   = "Errors"
    arrLogFormat(ARRAY_MASTER, COL_NOTES)                   = "Notes"
    arrLogFormat(ARRAY_MASTER, COL_METADATASTATE)           = "MetadataState"
    arrLogFormat(ARRAY_MASTER, COL_ISOFFICEPRODUCT)         = "IsOfficeProduct"
    arrLogFormat(ARRAY_MASTER, COL_PATCHFAMILY)             = "PatchFamily"
    arrLogFormat(ARRAY_MASTER, COL_OSPPLICENSE)             = "OSPP License"
    arrLogFormat(ARRAY_MASTER, COL_OSPPLICENSEXML)          = "OSPP License XML"
    arrLogFormat(ARRAY_MASTER, COL_ACIDLIST)                = "ACIDList"
    arrLogFormat(ARRAY_MASTER, COL_LICENSELIST)             = "Possible Licenses"
    arrLogFormat(ARRAY_MASTER, COL_UPGRADECODE)             = "UpgradeCode"
    arrLogFormat(ARRAY_MASTER, COL_VIRTUALIZED)             = "Virtualized"
    arrLogFormat(ARRAY_MASTER, COL_INSTALLTYPE)             = "InstallType"
    arrLogFormat(ARRAY_MASTER, COL_KEYCOMPONENTS)           = "KeyComponents"
    arrLogFormat(ARRAY_MASTER, COL_PRIMARYPRODUCTID)        = "Primary Product Id"

Const ARRAY_PATCH                                           = 1 'Patch data array id
    Const PATCH_COLUMNCOUNT                                 = 14
    Const PATCH_LOGSTART                                    = 1
    Const PATCH_LOGCHAINEDMAX                               = 8
    Const PATCH_LOGMAX                                      = 11
    Const PATCH_PRODUCT                                     = 0
    Const PATCH_KB                                          = 1
    Const PATCH_PACKAGE                                     = 3 ' PackageName
    Const PATCH_PATCHSTATE                                  = 2
    Const PATCH_SEQUENCE                                    = 4
    Const PATCH_UNINSTALLABLE                               = 5
    Const PATCH_INSTALLDATE                                 = 6
    Const PATCH_PATCHCODE                                   = 7
    Const PATCH_LOCALPACKAGE                                = 8
    Const PATCH_TRANSFORM                                   = 9
    Const PATCH_DISPLAYNAME                                 = 10
    Const PATCH_MOREINFOURL                                 = 11
    Const PATCH_CSP                                         = 12 ' Client side patch or patched AIP
    Const PATCH_CPOK                                        = 13 ' Local .msp package OK/available
    arrLogFormat(ARRAY_PATCH, PATCH_PRODUCT)                = "Patched Product: "
    arrLogFormat(ARRAY_PATCH, PATCH_KB)                     = "KB: "
    arrLogFormat(ARRAY_PATCH, PATCH_PACKAGE)                = "Package: "
    arrLogFormat(ARRAY_PATCH, PATCH_PATCHSTATE)             = "State: "
    arrLogFormat(ARRAY_PATCH, PATCH_SEQUENCE)               = "Sequence: "
    arrLogFormat(ARRAY_PATCH, PATCH_UNINSTALLABLE)          = "Uninstallable: "
    arrLogFormat(ARRAY_PATCH, PATCH_INSTALLDATE)            = "InstallDate: "
    arrLogFormat(ARRAY_PATCH, PATCH_PATCHCODE)              = "PatchCode: "
    arrLogFormat(ARRAY_PATCH, PATCH_LOCALPACKAGE)           = "LocalPackage: "
    arrLogFormat(ARRAY_PATCH, PATCH_TRANSFORM)              = "PatchTransform: "
    arrLogFormat(ARRAY_PATCH, PATCH_DISPLAYNAME)            = "DisplayName: "
    arrLogFormat(ARRAY_PATCH, PATCH_MOREINFOURL)            = "MoreInfoUrl: "
    arrLogFormat(ARRAY_PATCH, PATCH_CSP)                    = "ClientSidePatch: "
    arrLogFormat(ARRAY_PATCH, PATCH_CPOK)                   = "CachedMspOK: "

' arrMsp(MSPDEFAULT, MSP_COLUMNCOUNT)
Const ARRAY_MSPFILES                                        = 10
    Const MSPFILES_COLUMNCOUNT                              = 18
    Const MSPFILES_LOGMAX                                   = 9
    Const MSPFILES_TARGETS                                  = 0 ' Product Targets
    Const MSPFILES_KB                                       = 1
    Const MSPFILES_PACKAGE                                  = 2 ' PackageName
    Const MSPFILES_FAMILY                                   = 3
    Const MSPFILES_SEQUENCE                                 = 4
    Const MSPFILES_PATCHSTATE                               = 5
    Const MSPFILES_UNINSTALLABLE                            = 6
    Const MSPFILES_INSTALLDATE                              = 7
    Const MSPFILES_DISPLAYNAME                              = 8
    Const MSPFILES_MOREINFOURL                              = 9 
    Const MSPFILES_PATCHCODE                                = 10 
    Const MSPFILES_LOCALPACKAGE                             = 11 
    Const MSPFILES_BUCKET                                   = 12 
    Const MSPFILES_ATTRIBUTE                                = 13 ' Attribute msidbPatchSequenceSupersedeEarlier
    Const MSPFILES_TRANSFORM                                = 14 ' PatchTransform
    Const MSPFILES_XML                                      = 15 ' PatchXml
    Const MSPFILES_TABLES                                   = 16 ' PatchTables
    Const MSPFILES_CPOK                                     = 17 ' Local .msp package OK/available
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_TARGETS)          = "Patch Targets: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_KB)               = "KB: "
    arrLogFormat(ARRAY_MSPFILES , MSPFILES_PACKAGE)          = "Package: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_FAMILY)           = "Family: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_SEQUENCE)         = "Sequence: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_PATCHSTATE)       = "PatchState: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_UNINSTALLABLE)    = "Uninstallable: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_INSTALLDATE)      = "InstallDate: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_DISPLAYNAME)      = "DisplayName: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_MOREINFOURL)      = "MoreInfoUrl: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_PATCHCODE)        = "PatchCode: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_LOCALPACKAGE)     = "LocalPackage: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_BUCKET)           = "Bucket: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_ATTRIBUTE)        = "msidbPatchSequenceSupersedeEarlier: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_TRANSFORM)        = "PatchTransform: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_XML)              = "PatchXml: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_TABLES)           = "PatchTables: "
    arrLogFormat(ARRAY_MSPFILES, MSPFILES_CPOK)             = "CachedMspOK: "

Const ARRAY_AIPPATCH                                        = 11
    Const AIPPATCH_COLUMNCOUNT                              = 3
    Const AIPPATCH_PRODUCT                                  = 0
    Const AIPPATCH_PATCHCODE                                = 1
    Const AIPPATCH_DISPLAYNAME                              = 3
    arrLogFormat(ARRAY_AIPPATCH, AIPPATCH_PRODUCT)          = "Patched Product: "
    arrLogFormat(ARRAY_AIPPATCH, AIPPATCH_PATCHCODE)        = "PatchCode: "
    arrLogFormat(ARRAY_AIPPATCH, AIPPATCH_DISPLAYNAME)      = "DisplayName: "

Const ARRAY_FEATURE                                         = 2 'Feature data array id
    Const FEATURE_COLUMNCOUNT                               = 1
    Const FEATURE_PRODUCTCODE                               = 0
    Const FEATURE_TREE                                      = 1

Const ARRAY_ARP                                             = 4 'Add/remove products data array id
    Const ARP_CHILDOFFSET                                   = 6
    Const ARP_CONFIGPRODUCTCODE                             = 0
    Const COL_CONFIGNAME                                    = 1
    Const ARP_PRODUCTVERSION                                = 2
    Const COL_CONFIGINSTALLTYPE                             = 3
    Const COL_CONFIGPACKAGEID                               = 4
    Const COL_ARPALLPRODUCTS                                = 5
    Const COL_LBOUNDCHAINLIST                               = 6
    arrLogFormat(ARRAY_ARP, ARP_CONFIGPRODUCTCODE)          = "Config ProductCode"
    arrLogFormat(ARRAY_ARP, COL_CONFIGNAME)                 = "Config ProductName"
    arrLogFormat(ARRAY_ARP, ARP_PRODUCTVERSION)             = "ProductVersion"
    arrLogFormat(ARRAY_ARP, COL_CONFIGINSTALLTYPE)          = "Config InstallType"
    arrLogFormat(ARRAY_ARP, COL_CONFIGPACKAGEID)            = "Config PackageID"

Const ARRAY_IS                                              = 5 ' MSI InstallSource data array id
    Const UBOUND_IS                                         = 6
    Const IS_LOG_LBOUND                                     = 2
    Const IS_LOG_UBOUND                                     = 6
    Const IS_PRODUCTCODE                                    = 0
    Const IS_SOURCETYPE                                     = 1
    Const IS_SOURCETYPESTRING                               = 2
    Const IS_ORIGINALSOURCE                                 = 3
    Const IS_LASTUSEDSOURCE                                 = 4
    Const IS_LISRESILIENCY                                  = 5
    Const IS_ADDITIONALSOURCES                              = 6
    arrLogFormat(ARRAY_IS, IS_SOURCETYPESTRING)             = "InstallSource Type"
    arrLogFormat(ARRAY_IS, IS_ORIGINALSOURCE)               = "Initially Used Source"
    arrLogFormat(ARRAY_IS, IS_LASTUSEDSOURCE)               = "Last Used Source"
    arrLogFormat(ARRAY_IS, IS_LISRESILIENCY)                = "LIS Resiliency Sources"
    arrLogFormat(ARRAY_IS, IS_ADDITIONALSOURCES)            = "Network Sources"

Const ARRAY_VIRTPROD                                        = 6 ' Non MSI based virtualized products
    Const UBOUND_VIRTPROD                                   = 14
    Const VIRTPROD_KEYNAME                                  = 2
    Const VIRTPROD_CONFIGNAME                               = 3
    Const VIRTPROD_PRODUCTVERSION                           = 4
    Const VIRTPROD_SPLEVEL                                  = 5
    Const VIRTPROD_ISPROOFINGTOOLS                          = 6
    Const VIRTPROD_ISLANGUAGEPACK                           = 7
    Const VIRTPROD_OSPPLICENSE                              = 8
    Const VIRTPROD_OSPPLICENSEXML                           = 9
    Const VIRTPROD_ACIDLIST                                 = 10
    Const VIRTPROD_LICENSELIST                              = 11
    Const VIRTPROD_CHILDPACKAGES                            = 12
    Const VIRTPROD_KEYCOMPONENTS                            = 13
    Const VIRTPROD_EXCLUDEAPP                               = 14
    arrLogFormat(ARRAY_VIRTPROD, COL_PRODUCTCODE)           = "ProductCode"
    arrLogFormat(ARRAY_VIRTPROD, COL_PRODUCTNAME)           = "ProductName"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_KEYNAME)          = "KeyName"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME)       = "Config ProductName"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_SPLEVEL)          = "Version"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION)   = "BuildNumber"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_OSPPLICENSE)      = "(O)SPP License"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_OSPPLICENSEXML)   = "(O)SPP License XML"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_ACIDLIST)         = "ACIDList"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_LICENSELIST)      = "Possible Licenses"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CHILDPACKAGES)    = "Child Packages"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_KEYCOMPONENTS)    = "KeyComponents"
    arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_EXCLUDEAPP)       = "Excluded Applications"

'License Data Array
Const ARRAY_LICENSE                   = 7 
    Const UBOUND_LICENSE              = 10
    Const LICENSE_ACID          = 0
    arrLogFormat(ARRAY_LICENSE, LICENSE_ACID) = "ACID"


Const CSV                             = ", "
Const DSV                             = " - "
Const DOT                             = ". "
Const ERR_CATEGORYNOTE                = "Note: "
Const ERR_CATEGORYWARN                = "Warning: "
Const ERR_CATEGORYERROR               = "Error: "
Const ERR_NONADMIN                    = "The script appears to run outside administrator context"
Const ERR_NONELEVATED                 = "The script does not appear to run elevated"
Const ERR_DATAINTEGRITY               = "A script internal error occurred. The integrity of the logged data might be affected"
Const ERR_OBJPRODUCTINFO              = "Installer.ProductInfo -> "
Const ERR_INITSUMINFO                 = "Could not connect to summary information stream"
Const ERR_NOARRAY                     = "Array check failed"
Const ERR_UNKNOWNHANDLER              = "Unknown Error Handler: '"
Const ERR_PRODUCTSEXALL               = "ProductsEx for MSIINSTALLCONTEXT_ALL failed"
Const ERR_PATCHESEX                   = "PatchesEx failed to get a list of patches for: "
Const ERR_PATCHES                     = "Installer.Patches failed to get a list of patches"
Const ERR_MISSINGCHILD                = "A chained product is missing which breaks the ability to maintain or uninstall this product. "
Const ERR_ORPHANEDITEM                = "Office application without entry point in Add/Remove Programs"
Const ERR_INVALIDPRODUCTCODE          = "Critical Windows Installer metadata corruption detected 'Invalid ProductCode'"
Const ERR_INVALIDGUID                 = "GUID validation failed"
Const ERR_INVALIDGUIDCHAR             = "Guid contains invalid character(s)"
Const ERR_INVALIDGUIDLENGTH           = "Invalid length for GUID  "
Const ERR_GUIDCASE                    = "Guid contains lower case character(s)"
Const ERR_BADARPMETADATA              = "Crititcal ARP metadata corruption detected in key: "
Const ERR_OFFSCRUB_TERMINATED         = "Bad ARP metadata. This can be caused by an OffScrub run that was terminated before it could complete:"
Const ERR_ARPENTRYMISSING             = "Expected regkey not present for ARP config parent"
Const ERR_REGKEYMISSING               = "Regkey does not exist: "
Const ERR_CUSTOMSTACKCORRUPTION       = "Custom stack list string corrupted"
Const ERR_BADMSPMETADATA              = "Metadata mismatch for patch registration"
Const ERR_BADMSINAMEMETADATA          = "Failed to retrieve value for original .msi name"
Const ERR_BADPACKAGEMETADATA          = "Failed to retrieve value for cached .msi package"
Const ERR_PACKAGEAPIFAILURE           = "API failed to retrieve value for cached .msi package"
Const ERR_BADPACKAGECODEMETADATA      = "Failed to retrieve value for Package Code"
Const ERR_PACKAGECODEMISMATCH         = "PackageCode mismatch between registered value and cached .msi"
Const ERR_LOCALPACKAGEMISSING         = "Local cached .msi appears to be missing"
Const ERR_BADTRANSFORMSMETADATA       = "Failed to retrieve value for Transforms"
Const ERR_SICONNECTFAILED             = "Failed to connect to SummaryInformation stream"
Const ERR_MSPOPENFAILED               = "OpenDatabase failed to open .msp file "
Const ERR_MSIOPENFAILED               = "OpenDatabase failed to open .msi file "
Const ERR_BADFILESTATE                = " has unexpected file state(s). "
Const ERR_FILEVERSIONLOW              = "Review file versions for product "
Const BPA_GUID                        = "For details on 'GUID' see https://msdn.microsoft.com/en-us/library/Aa368767.aspx"
Const BPA_PACKAGECODE                 = "For details on 'Package Codes' see https://msdn.microsoft.com/en-us/library/aa370568.aspx"
Const BPA_PRODUCTCODE                 = "For details on 'Product Codes' see https://msdn.microsoft.com/en-us/library/aa370860.aspx"
Const BPA_PACKAGECODEMISMATCH         = "A mismatch of the PackageCode will force the Windows Installer to recache the local .msi from the InstallSource. For details on 'Package Code' see https://msdn.microsoft.com/en-us/library/aa370568.aspx"
'=======================================================================================================

Main

'=======================================================================================================
'Module Main
'=======================================================================================================
Sub Main
    Dim fCheckPreReq, FsoLogFile, FsoXmlLogFile
    On Error Resume Next
' Check type of scripting host
    fCScript = (UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C")
' Ensure all required objects are available. Prerequisite Checks with inline error handling. 
    fCheckPreReq = CheckPreReq()
    If fCheckPreReq = False Then Exit Sub
' Initializations 
    Initialize 
' Get computer specific properties
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 1 of 11: ComputerProperties"
    ComputerProperties 
' Build an array with a list of all Windows Installer based products
' After this point the master array "arrMaster" is instantiated and has the basic product details
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 2 of 11: Product detection"
    FindAllProducts 
' Get additional product properties and add them to the master array
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 3 of 11: ProductProperties"
    ProductProperties 
' Build an array with data on the InstallSource(s)
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 4 of 11: InstallSources"
    ReadMsiInstallSources
' Build an array with data from Add/Remove Products
' Only Office >= 2007 products that use a multiple .msi structure will be covered here
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 5 of 11: Add/Remove Programs analysis"
    ARPData 
' Add Licensing data
' Only Office >= 2010 products that use OSPP are covered here
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 6 of 11: Licensing (OSPP)"
    OsppCollect 
' Build an array with all patch data.
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 7 of 11: Patch detection"
    FindAllPatches
    If fFeatureTree Then
    ' Build a tree structure for the Features 
        If fCScript AND NOT fQuiet Then wscript.echo "Stage 8 of 11: FeatureStates"
        FindFeatureStates
    Else
        If fCScript AND NOT fQuiet Then wscript.echo "Skipping stage 8 of 11: (FeatureStates)"
    End If 'fFeatureTree
' Create file inventory XML files
    If fFileInventory Then
        If fCScript AND NOT fQuiet Then wscript.echo "Stage 9 of 11: FileInventory"
        FileInventory
    Else
        If fCScript AND NOT fQuiet Then wscript.echo "Skipping stage 9 of 11: (FileInventory)"
    End If 'fFileInventory
' Prepare the collected data for the output file
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 10 of 11: Prepare collected data for output"
    PrepareLog sLogFormat 
' Write the output file
    If fCScript AND NOT fQuiet Then wscript.echo "Stage 11 of 11: Write log"
    WriteLog 
    Set FsoLogFile = oFso.GetFile(sLogFile)
    Set FsoXmlLogFile = oFso.GetFile(sPathOutputFolder & sComputerName & "_ROIScan.xml")
    If fFileInventory Then 
        If (oFso.FileExists(sPathOutputFolder & sComputerName & "_ROIScan.zip") AND NOT fZipError) Then
            CopyToZip ShellApp.NameSpace(sPathOutputFolder & sComputerName & "_ROIScan.zip"), FsoLogFile
            CopyToZip ShellApp.NameSpace(sPathOutputFolder & sComputerName & "_ROIScan.zip"), FsoXmlLogFile
        End If
    End If
    If fCScript AND NOT fQuiet Then wscript.echo "Done!"
' Open the output file
    If Not fQuiet Then
        Set oShell = CreateObject("WScript.Shell")
        If fFileInventory Then 
            If oFso.FileExists(sPathOutputFolder & sComputerName & "_ROIScan.zip") AND NOT fZipError Then
                oShell.Run "explorer /e," & chr(34) & sPathOutputFolder & sComputerName & "_ROIScan.zip" & chr(34) 
            Else 
                oShell.Run "explorer /e," & chr(34) & sPathOutputFolder & "ROIScan" & chr(34)
            End If
        End If 'fFileInventory
        oShell.Run chr(34) & sLogFile & chr(34)
        Set oShell = Nothing
    End If 'fQuiet
' Clear up Objects
    CleanUp
End Sub
'=======================================================================================================

'Initialize defaults, setting and collect current user information
Sub Initialize
    Dim oApp, Process, Processes
    Dim sEnvVar, Argument
    Dim iPopup, iInstanceCnt
    Dim fPrompt
    On Error Resume Next
    'Ensure there's only a single instance running of this script
    iInstanceCnt = 0
    wscript.sleep 500
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If LCase(Mid(Process.Name, 2, 6)) = "script" Then 
            If InStr(LCase(Process.CommandLine), "roiscan") > 0 AND NOT InStr(Process.CommandLine, " UAC") > 0 Then iInstanceCnt = iInstanceCnt + 1
        End If
    Next 'Process
    If iInstanceCnt > 1 Then
        If NOT fQuiet Then wscript.echo ERR_CATEGORYERROR & "Another instance of this script is already running."
        wscript.quit
    End If
   
    'Other defaults
    Set dicPatchLevel = CreateObject("Scripting.Dictionary")
    Set dicScenario = CreateObject("Scripting.Dictionary")
    Set dicC2RPropV2 = CreateObject("Scripting.Dictionary")
    Set dicScenarioV2 = CreateObject("Scripting.Dictionary")
    Set dicKeyComponentsV2 = CreateObject("Scripting.Dictionary")
    Set dicVirt2Cultures = CreateObject("Scripting.Dictionary")
    Set dicC2RPropV3 = CreateObject("Scripting.Dictionary")
    Set dicScenarioV3 = CreateObject("Scripting.Dictionary")
    Set dicKeyComponentsV3 = CreateObject("Scripting.Dictionary")
    Set dicVirt3Cultures = CreateObject("Scripting.Dictionary")
    Set dicProductCodeC2R = CreateObject("Scripting.Dictionary")
    Set dicMapArpToConfigID = CreateObject("Scripting.Dictionary")
    Set dicArp = CreateObject("Scripting.Dictionary")
    Set dicPolHKCU = CreateObject("Scripting.Dictionary")
    Set dicPolHKLM = CreateObject("Scripting.Dictionary")
    fZipError = False
    fInitArrProdVer = False

    ' log output folder
    If sPathOutputFolder = "" Then sPathOutputFolder = "%TEMP%"
    sPathOutputFolder = oShell.ExpandEnvironmentStrings(sPathOutputFolder)
    sPathOutputFolder = Replace(sPathOutputFolder, "'","")
    If sPathOutputFolder = "." OR Left(sPathOutputFolder, 2) = ".\" Then sPathOutputFolder = GetFullPathFromRelative(sPathOutputFolder)
    sPathOutputFolder = oFso.GetAbsolutePathName(sPathOutputFolder)
    If Trim(UCase(sPathOutputFolder)) = "DESKTOP" Then
        Set oApp = CreateObject ("Shell.Application")
        Const DESKTOP = &H10&
        sPathOutputFolder = oApp.Namespace(DESKTOP).Self.Path
        Set oApp = Nothing
    End If
    If Not oFso.FolderExists(sPathOutputFolder) Then 
        ' custom log folder location does not exist.
        ' try to create the folder before falling back to default
        oFso.CreateFolder sPathOutputFolder
        If NOT Err = 0 Then
            sPathOutputFolder = oShell.ExpandEnvironmentStrings("%TEMP%") & "\" 
            Err.Clear
        End If
    End If
    While Right(sPathOutputFolder, 1) = "\" 
        sPathOutputFolder = Left(sPathOutputFolder, Len(sPathOutputFolder) - 1)
    Wend
    If Not Right(sPathOutputFolder, 1) = "\" Then sPathOutputFolder = sPathOutputFolder & "\"
    sLogFile = sPathOutputFolder & sComputerName & "_ROIScan.log"

    CacheLog LOGPOS_COMPUTER, LOGHEADING_H1, Null, "Computer" 
    CacheLog LOGPOS_REVITEM, LOGHEADING_H1, Null, "Review Items" 
    CacheLog LOGPOS_RAW, LOGHEADING_H1, Null, "Raw Data" 
    iPopup = -1
    fPrompt = True
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            If Argument = "UAC" Then fPrompt = False
        Next 'Argument
    End If
'Add warning to log if non-admin was detected
    If Not fIsAdmin Then 
        Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_NONADMIN
        If NOT fQuiet AND fPrompt Then RelaunchElevated
    End If
    If fIsAdmin AND (NOT fIsElevated) Then 
        Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_NONELEVATED
        If NOT fQuiet AND fPrompt Then RelaunchElevated
   End If
'Ensure CScript as engine
    If (NOT UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "C") AND (NOT fDisallowCScript) Then RelaunchAsCScript
'Check on 64 bit OS -> see CheckPreReq
'Init sCurUserSid
    GetUserSids("Current") 
'Init "arrUUSids"
    Redim arrUUSids(-1)
    GetUserSids ("UserUnmanaged") 
'Init "arrUMSids"
    Redim arrUMSids(-1)
    GetUserSids("UserManaged") 
'Init KeyComponents dictionary
    Set dicKeyComponents = CreateObject("Scripting.Dictionary")
    InitKeyComponents
    iVMVirtOverride = 0
'Set defaults for ProductList arrays
    InitPLArrays 
    Set dicMspIndex = CreateObject("Scripting.Dictionary")
    bOsppInit = False
End Sub 'Initialize
'=======================================================================================================
'End Of Main Module

'=======================================================================================================
'Module FileInventory
'=======================================================================================================

'File inventory for installed applications
Sub FileInventory
    Dim sProductCode, sQueryFT, sQueryMst, sQueryCompID, sQueryFC, sMst, sFtk, sCmp
    Dim sPath, sName, sXmlLine, sAscCheck, sQueryDir, sQueryAssembly, sPatchCode
    Dim sCurVer, sSqlCreateTable, sTables, sTargetPath
    Dim iPosMaster, iPosPatch, iFoo, iFile, iMst, iCnt, iAsc, iAscCnt, iCmp
    Dim iPosArr, iArrMaxCnt, iColCnt, iBaseRefCnt, iIndex
    Dim bAsc, bMstApplied, bBaseRefFound, bFtkViewComplete, bFtkInScope, bFtkForceOutOfScope, bNeedKeyPathFallback
    Dim bCreate, bDelete, bDrop, bInsert, bFileNameChanged
    Dim MsiDb, SessionDb, MspDb, qViewFT, qViewMst, qViewCompID, qViewFC, qViewDir, qViewAssembly, Record, Record2, Record3
    Dim qViewMspCompId, qViewMspFC, qViewMsiAssembly, AllOfficeFiles, FileStream
    Dim SessionDir, Table, ViewTables, tbl
    Dim dicTransforms, dicKeys
    Dim arrTables
    
    If fBasicMode Then Exit Sub
    On Error Resume Next
    Const FILES_FTK             = 0 'key field for each file
    Const FILES_SOURCE          = 1
    Const FILES_ISPATCHED       = 2
    Const FILES_FILESTATUS      = 3 
    Const FILES_FILE            = 4
    Const FILES_FOLDER          = 5
    Const FILES_FULLNAME        = 6
    Const FILES_LANGUAGE        = 7
    Const FILES_BASEVERSION     = 8
    Const FILES_PATCHVERSION    = 9
    Const FILES_CURRENTVERSION  = 10
    Const FILES_VERSIONSTATUS   = 11
    Const FILES_PATCHCODE       = 12 'key field for patch
    Const FILES_PATCHSTATE      = 13
    Const FILES_PATCHKB         = 14
    Const FILES_PATCHPACKAGE    = 15
    Const FILES_PATCHMOREINFO   = 16
    Const FILES_DIRECTORY       = 17
    Const FILES_COMPONENTID     = 18
    Const FILES_COMPONENTNAME   = 19
    Const FILES_COMPONENTSTATE  = 20
    Const FILES_KEYPATH         = 21
    Const FILES_COMPONENTCLIENTS= 22
    Const FILES_FEATURENAMES    = 23
    Const FILES_COLUMNCNT       = 23
    
    Const SQL_FILETABLE = "SELECT * FROM `_TransformView` WHERE `Table` = 'File' ORDER BY `Row`"
    Const SQL_PATCHTRANSFORMS = "SELECT `Name` FROM `_Storages` ORDER BY `Name`"
    Const INSTALLSTATE_ASSEMBLY = 6
    Const WAITTIME = 500 

' Loop all products
    If fCScript AND NOT fQuiet Then wscript.echo vbTab & "File version scan"
    For iPosMaster = 0 To UBound(arrMaster)
        If (arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)) AND (arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
        ' Cache ProductCode
            sProductCode = "" : sProductCode = arrMaster(iPosMaster, COL_PRODUCTCODE)
            If fCScript AND NOT fQuiet Then wscript.echo vbTab & sProductCode 
        ' Reset Files array
            iPosArr = -1
            iArrMaxCnt = 5000
            ReDim arrFiles(FILES_COLUMNCNT, iArrMaxCnt)
            iBaseRefCnt = 0
            For iFoo = 1 To 1
            ' Add fields from msi base
            ' ------------------------
            ' Connect to the local .msi file for reading
                Err.Clear
                Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster, COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
                If Not Err = 0 Then 
                    Exit For
                End If 'Err = 0
            ' Check which tables exist in the current .msi
                sTables = ""
                Set ViewTables = MsiDb.OpenView("SELECT `Name` FROM `_Tables` ORDER BY `Name`")
                ViewTables.Execute
                Do
                    Set Table = ViewTables.Fetch
                    If Table Is Nothing then Exit Do
                    sTables = sTables & Table.StringData(1) & ","
                    If Not Err = 0 Then Exit Do
                Loop
                ViewTables.Close
                arrTables = Split(RTrimComma(sTables), ",")
            ' Build an assembly reference dictionary
                Set dicAssembly = Nothing
                Set dicAssembly = CreateObject("Scripting.Dictionary")
                If InStr(sTables, "MsiAssembly,") > 0 Then
                    sQueryAssembly = "SELECT DISTINCT `Component_` FROM MsiAssembly"
                    Set qViewAssembly = MsiDb.OpenView(sQueryAssembly)
                    qViewAssembly.Execute
                ' If the MsiAssmbly table does not exist it returns an error 
                    If Not Err = 0 Then Err.Clear
                    Set Record = qViewAssembly.Fetch
                ' must not enter the loop in case of an error!
                    If Not Err = 0 Then
                        Err.Clear
                    Else 
                        Do Until Record Is Nothing
                            If Not dicAssembly.Exists(Record.StringData(1)) Then
                                dicAssembly.Add Record.StringData(1), Record.StringData(1)
                            End If
                            Set Record = qViewAssembly.Fetch
                        Loop
                    End If 'Not Err = 0
                    qViewAssembly.Close
                End If 'InStr(sTables, "MsiAssembly") > 0
                
                If InStr(sTables, "SxsMsmGenComponents,") > 0 Then
                    sQueryAssembly = "SELECT DISTINCT `Component_` FROM SxsMsmGenComponents"
                    Set qViewAssembly = MsiDb.OpenView(sQueryAssembly)
                    qViewAssembly.Execute
                ' If the MsiAssmbly table does not exist it returns an error 
                    If Not Err = 0 Then Err.Clear
                    Set Record = qViewAssembly.Fetch
                ' must not enter the loop in case of an error!
                    If Not Err = 0 Then
                        Err.Clear
                    Else 
                        Do Until Record Is Nothing
                            If Not dicAssembly.Exists(Record.StringData(1)) Then
                                dicAssembly.Add Record.StringData(1), Record.StringData(1)
                            End If
                            Set Record = qViewAssembly.Fetch
                        Loop
                    End If 'Not Err = 0
                    qViewAssembly.Close
                End If 'InStr(sTables, "MsiAssembly") > 0
                
            ' Build directory reference
                Set SessionDir = Nothing
                oMsi.UILevel = 2 'None
                Set SessionDir = oMsi.OpenProduct(sProductCode)
                SessionDir.DoAction("CostInitialize")
                SessionDir.DoAction("FileCost")
                SessionDir.DoAction("CostFinalize")
                Set dicFolders = Nothing
                Set dicFolders = CreateObject("Scripting.Dictionary")
                Err.Clear
                Set SessionDb = SessionDir.Database
                sQueryDir = "SELECT DISTINCT `Directory` FROM Directory"
                Set qViewDir = SessionDb.OpenView(sQueryDir)
                qViewDir.Execute
                Set Record = qViewDir.Fetch
            ' must not enter the loop in case of an error!
                If Not Err = 0 Then
                    Err.Clear
                Else 
                    Do Until Record Is Nothing
                        If Not dicFolders.Exists(Record.Stringdata(1)) Then
                            sTargetPath = "" : sTargetPath = SessionDir.TargetPath(Record.Stringdata(1))
                            If NOT sTargetPath = "" Then dicFolders.Add Record.Stringdata(1), sTargetPath
                        End If
                        Set Record = qViewDir.Fetch
                    Loop
                End If 'Not Err = 0
                qViewDir.Close
            
            ' .msi file inventory
            ' -------------------
                sQueryFT = "SELECT * FROM File"
                Set qViewFT = MsiDb.OpenView(sQueryFT)
                qViewFT.Execute
                Set Record = qViewFT.Fetch()
                Do Until Record Is Nothing
                ' Next Row in Array
                ' -----------------
                    iPosArr = iPosArr + 1
                    iBaseRefCnt = iPosArr
                    If iPosArr > iArrMaxCnt Then
                    ' increase array row buffer
                        iArrMaxCnt = iArrMaxCnt + 1000
                        ReDim Preserve arrFiles(FILES_COLUMNCNT, iArrMaxCnt)
                    End If 'iPosArr > iArrMaxCnt
                ' add FTK name
                    arrFiles(FILES_FTK, iPosArr) = Record.StringData(1)
                ' the FilesSource flag allows to filter the data in the report to exclude patch only entries.
                    arrFiles(FILES_SOURCE, iPosArr) = "Msi"
                ' default IsPatched field to 'False'
                    arrFiles(FILES_ISPATCHED, iPosArr) = False
                ' add the LFN (long filename)
                    arrFiles(FILES_FILE, iPosArr) = GetLongFileName(Record.StringData(3))
                ' add ComponentName
                    arrFiles(FILES_COMPONENTNAME, iPosArr) = Record.StringData(2)
                ' add ComponentID and Directory reference from Component table
                    sQueryCompID = "SELECT `Component`, `ComponentId`, `Directory_` FROM Component WHERE `Component` = '" & Record.StringData(2) & "'"
                    Set qViewCompID = MsiDb.OpenView(sQueryCompID)
                    qViewCompID.Execute
                    Set Record2 = qViewCompID.Fetch()
                    arrFiles(FILES_COMPONENTID, iPosArr) = Record2.StringData(2)
                    arrFiles(FILES_DIRECTORY, iPosArr) = Record2.StringData(3)
                    Set Record2 = Nothing
                    qViewCompID.Close
                    Set qViewCompID = Nothing
                ' ComponentState
                    arrFiles(FILES_COMPONENTSTATE, iPosArr) = GetComponentState(sProductCode, arrFiles(FILES_COMPONENTID, iPosArr), iPosMaster)
                ' add ComponentClients
                    arrFiles(FILES_COMPONENTCLIENTS, iPosArr) = GetComponentClients(arrFiles(FILES_COMPONENTID, iPosArr), arrFiles(FILES_COMPONENTSTATE, iPosArr))
                ' add Features that use the component
                    sQueryFC = "SELECT * FROM FeatureComponents WHERE `Component_` = '" & Record.StringData(2)  & "'"
                    Set qViewFC = MsiDb.OpenView(sQueryFC)
                    qViewFC.Execute
                    Set Record2 = qViewFC.Fetch()
                    Do Until Record2 Is Nothing
                        arrFiles(FILES_FEATURENAMES, iPosArr) = arrFiles(FILES_FEATURENAMES, iPosArr) & Record2.StringData(1) & _
                                                               "(" & TranslateFeatureState(oMsi.FeatureState(sProductCode, Record2.StringData(1))) & ")" &  ","
                        Set Record2 = qViewFC.Fetch()
                    Loop
                    RTrimComma arrFiles(FILES_FEATURENAMES, iPosArr)
                ' add KeyPath
                    arrFiles(FILES_KEYPATH, iPosArr) = GetComponentPath(sProductCode, arrFiles(FILES_COMPONENTID, iPosArr), arrFiles(FILES_COMPONENTSTATE, iPosArr))
                    sPath = "" : sName = ""
                ' add Componentpath
                    If dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME, iPosArr)) Then
                    ' Assembly
                        If arrFiles(FILES_COMPONENTSTATE, iPosArr) = INSTALLSTATE_LOCAL Then
                            sPath = GetAssemblyPath(arrFiles(FILES_FILE, iPosArr), arrFiles(FILES_KEYPATH, iPosArr), dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr)))
                            arrFiles(FILES_FOLDER, iPosArr) = Left(sPath, InStrRev(sPath, "\"))
                        Else
                            arrFiles(FILES_FOLDER, iPosArr) = sPath
                        End If
                    Else
                    ' Regular component
                        arrFiles(FILES_FOLDER, iPosArr) = dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr))
                        If arrFiles(FILES_FOLDER, iPosArr) = "" Then
                            ' failed to obtain the directory from the session object
                            ' try again by direct read from session object
                            arrFiles(FILES_FOLDER, iPosArr) = SessionDir.TargetPath(arrFiles(FILES_DIRECTORY, iPosArr))

                            ' if still failed, fall back to the keypath by using the assembly logic to resolve the path
                            If arrFiles(FILES_FOLDER, iPosArr) = "" AND arrFiles(FILES_COMPONENTSTATE, iPosArr) = INSTALLSTATE_LOCAL Then
                                sPath = GetAssemblyPath(arrFiles(FILES_FILE, iPosArr), arrFiles(FILES_KEYPATH, iPosArr), dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr)))
                                arrFiles(FILES_FOLDER, iPosArr) = Left(sPath, InStrRev(sPath, "\"))
                            End If
                        End If
                    End If
                ' add file FullName - if sPath contains a string then it's the result of the assembly detection
                    If sPath = "" Then
                        arrFiles(FILES_FULLNAME, iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FOLDER, iPosArr), arrFiles(FILES_FILE, iPosArr))
                    Else
                        arrFiles(FILES_FULLNAME, iPosArr) = sPath
                        sName = Right(sPath, Len(sPath) -InStrRev(sPath, "\"))
                        'Update the files field 
                        If Not UCase(sName) = UCase(arrFiles(FILES_FILE, iPosArr)) Then _
                            arrFiles(FILES_FILE, iPosArr) = sName 
                    End If
                ' add (msi) BaseVersion
                    If Not Err = 0 Then Err.Clear
                    arrFiles(FILES_BASEVERSION, iPosArr) = Record.StringData(5)
                ' add FileState and FileVersion
                    arrFiles(FILES_FILESTATUS, iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                    arrFiles(FILES_CURRENTVERSION, iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                ' add Language
                    arrFiles(FILES_LANGUAGE, iPosArr) = Record.StringData(6)
                ' get next row
                    Set Record = qViewFT.Fetch()
                Loop
                Set Record = Nothing
                qViewFT.Close
                Set qViewFT = Nothing

            ' --------------
            '  Add patches '
            ' --------------
            ' loop through all patches for the current product
                For iPosPatch = 0 to UBound(arrPatch, 3)
                    If Not (IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))) AND (arrPatch(iPosMaster, PATCH_CSP, iPosPatch) = True) Then
                        Err.Clear
                        sPatchCode = arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)
                        Set MspDb = oMsi.OpenDatabase(arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iPosPatch), MSIOPENDATABASEMODE_PATCHFILE)
                        If Err = 0 Then
                        ' create the table structures from .msi schema to allow detailed query of the .msp _TransformView 
                            For Each tbl in arrTables
                                sSqlCreateTable = "CREATE TABLE `" & tbl & "` (" & GetTableColumnDef(MsiDb, tbl) & " PRIMARY KEY " & GetPrimaryTableKeys(MsiDb, tbl)  & ")"
                                MspDb.OpenView(sSqlCreateTable).Execute
                            Next 'tbl
                        ' check if a PatchTransform is available
                            sMst = "" : bMstApplied = False
                            sMst = arrPatch(iPosMaster, PATCH_TRANSFORM, iPosPatch)
                            Err.Clear
                            If InStr(sMst, ";") > 0 Then  
                                bMstApplied = True 
                                sMst = Left(sMst, InStr(sMst, ";") - 1) 
                            ' apply the patch transform  
                            ' msiTransformErrorAll includes msiTransformErrorViewTransform which creates the "_TransformView" table 
                                MspDb.ApplyTransform sMst, MSITRANSFORMERROR_ALL 
                            End If 'InStr(sMst, ";") > 0 
                        ' if no known .mst or failed to apply the .mst we go into generic patch embedded transform detection loop
                            If (Not bMstApplied) OR (Not Err = 0) Then
                                Err.Clear
                            ' Dictionary object for the patch transforms
                                Set dicTransforms = CreateObject("Scripting.Dictionary")
                            ' create the view to retrieve the patch transforms
                                sQueryMst = SQL_PATCHTRANSFORMS
                                Set qViewMst = MspDb.OpenView(sQueryMst): qViewMst.Execute
                                Set Record = qViewMst.Fetch
                            ' loop all transforms and add them to the dictionary
                                Do Until Record Is Nothing
                                    sMst = Record.StringData(1)
                                        dicTransforms.Add sMst, sMst
                                    Set Record = qViewMst.Fetch
                                Loop
                                qViewMst.Close : Set qViewMst = Nothing
                                Set Record = Nothing
                            ' apply the patch transforms
                                dicKeys = dicTransforms.Keys
                                For iMst = 0 To dicTransforms.Count - 1
                                ' get the transform name
                                    sMst = dicKeys(iMst)
                                ' apply the patch transform / staple them all on the table
                                    MspDb.ApplyTransform ":" & sMst, MSITRANSFORMERROR_ALL
                                    If Not Err = 0 Then Err.Clear
                                Next 'iMst
                            End If '(Not bMstApplied) OR (Not Err = 0)
                            
                        ' _TransformView loop
                        ' -------------------
                        ' update the MsiAssembly reference dictionary
                            Err.Clear
                            If InStr(sTables, "MsiAssembly,") > 0 Then
                                Set qViewMsiAssembly = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='MsiAssembly' ORDER By `Row`")
                                qViewMsiAssembly.Execute()
                                Set Record = qViewMsiAssembly.Fetch()
                                Do
                                    If Record Is Nothing Then Exit Do
                                    If Not Err = 0 Then
                                        Err.Clear
                                        Exit Do
                                    End If
                                    If Record.StringData(2) = "INSERT" Then
                                        If Not dicAssembly.Exists(Record.StringData(3)) Then dicAssembly.Add Record.StringData(3), Record.StringData(3)
                                    End If
                                    Set Record = qViewMsiAssembly.Fetch()
                                Loop
                                qViewMsiAssembly.Close
                            End If
                            Set Record = Nothing

                            If InStr(sTables, "SxsMsmGenComponents,") > 0 Then
                                Set qViewMsiAssembly = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='SxsMsmGenComponents' ORDER By `Row`")
                                qViewMsiAssembly.Execute()
                                Set Record = qViewMsiAssembly.Fetch()
                                Do
                                    If Record Is Nothing Then Exit Do
                                    If Not Err = 0 Then
                                        Err.Clear
                                        Exit Do
                                    End If
                                    If Record.StringData(2) = "INSERT" Then
                                        If Not dicAssembly.Exists(Record.StringData(3)) Then dicAssembly.Add Record.StringData(3), Record.StringData(3)
                                    End If
                                    Set Record = qViewMsiAssembly.Fetch()
                                Loop
                                qViewMsiAssembly.Close
                            End If
                            Set Record = Nothing
                        ' get the files being modified from the "_TransformView" 'File' table
                            Set qViewMst = MspDb.OpenView(SQL_FILETABLE) : qViewMst.Execute()
                        ' loop all of the entries in the File table from "_TransformView"
                            Set Record = qViewMst.Fetch()
                        ' initial defaults
                            sFtk = ""
                            bFtkViewComplete = True
                            bFtkInScope = False
                            Do
                            ' is this the next FTK?
                                If (Not sFtk = Record.StringData(3)) OR (Record Is Nothing) Then
                                    If Record Is Nothing Then Err.Clear
                                ' yes this is the next FTK or the last time before exit of the loop
                                ' is previous FTK handling complete?
                                    If Not bFtkViewComplete Then
                                    ' previous FTK handling is not complete
                                    ' is previous FTK in scope?
                                        If bFtkInScope AND NOT bFtkForceOutOfScope Then
                                        'FTK is in scope - reset the scope flag
                                            bFtkInScope = False
                                            If bBaseRefFound Then
                                            ' update base entry fields with patch information
                                            ' check if the filename got updated
                                                If bFileNameChanged Then
                                                    arrFiles(FILES_FILE, iPosArr) = GetLongFileName(arrFiles(FILES_FILE, iPosArr))
                                                ' the filename got changed by a patch
                                                ' if the patch is in the 'Applied' state -> care about this change
                                                    If LCase(arrFiles(FILES_PATCHSTATE, iPosArr)) = "applied" Then
                                                    ' update the filename in the baseref if this is (NOT Assembly) OR (broken)
                                                        If NOT(dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME, iPosArr))) OR (arrFiles(FILES_FILESTATUS, iPosArr) = INSTALLSTATE_BROKEN) Then
                                                        ' correct the baseref filename field
                                                            arrFiles(FILES_FILE, iCnt) = arrFiles(FILES_FILE, iPosArr)
                                                        ' File FullName
                                                            arrFiles(FILES_FULLNAME, iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FOLDER, iPosArr), arrFiles(FILES_FILE, iPosArr))
                                                        ' recheck the filestate
                                                            arrFiles(FILES_FILESTATUS, iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                                                            arrFiles(FILES_FILESTATUS, iCnt) = arrFiles(FILES_FILESTATUS, iPosArr)
                                                        ' recheck the version
                                                            arrFiles(FILES_CURRENTVERSION, iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                                                            arrFiles(FILES_CURRENTVERSION, iCnt) = arrFiles(FILES_CURRENTVERSION, iPosArr)
                                                        End If
                                                    End If 'applied
                                                End If 'bFileNameChanged
                                            ' set IsPatched flag
                                                arrFiles(FILES_ISPATCHED, iCnt) = True
                                            ' check if PatchVersion field needs to be updated
                                                iCmp = 2 : sCmp = ""
                                                iCmp = CompareVersion(arrFiles(FILES_PATCHVERSION, iPosArr), arrFiles(FILES_PATCHVERSION, iCnt), True)
                                                Select Case iCmp
                                                Case VERSIONCOMPARE_LOWER   ': sCmp= "ERROR_VersionLow"
                                                Case VERSIONCOMPARE_MATCH   ': sCmp= "SUCCESS_VersionMatch"
                                                Case VERSIONCOMPARE_HIGHER  ': sCmp= "SUCCESS_VersionHigh"
                                                ' update the base field
                                                    arrFiles(FILES_PATCHVERSION, iCnt) = arrFiles(FILES_PATCHVERSION, iPosArr)
                                                Case VERSIONCOMPARE_INVALID ': sCmp= ""
                                                End Select
                                            ' update PatchCode field
                                                arrFiles(FILES_PATCHCODE, iCnt) = arrFiles(FILES_PATCHCODE, iCnt) & sPatchCode & ","
                                            ' update PatchKB field 
                                                If dicMspIndex.Exists(sPatchCode) Then arrFiles(FILES_PATCHKB, iCnt) = arrFiles(FILES_PATCHKB, iCnt) & arrFiles(FILES_PATCHKB, iPosArr) & ","
                                            ' update PatchMoreInfo field
                                                arrFiles(FILES_PATCHMOREINFO, iCnt) = arrFiles(FILES_PATCHMOREINFO, iCnt) & arrPatch(iPosMaster, PATCH_MOREINFOURL, iPosPatch) & ","
                                            Else    
                                            ' bBaseRefFound is False. 
                                            ' -----------------------
                                            ' this is a new file introduced with the patch. set FileSource flag
                                                arrFiles(FILES_SOURCE, iPosArr) = "Msp"
                                            ' define IsPatched field
                                                arrFiles(FILES_ISPATCHED, iPosArr) = False
                                            ' add the LFN (long file name)
                                                arrFiles(FILES_FILE, iPosArr) = GetLongFileName(arrFiles(FILES_FILE, iPosArr))
                                            ' locate the ComponentId from Component table
                                                sQueryCompID = "SELECT `Component`, `ComponentId`, ´Directory_` FROM Component WHERE `Component` = '" & arrFiles(FILES_COMPONENTNAME, iPosArr) & "'"
                                                Set qViewCompID = MsiDb.OpenView(sQueryCompID)
                                                If Not Err = 0 Then 
                                                    Err.Clear
                                                    Set Record2 = Nothing
                                                Else
                                                    qViewCompID.Execute
                                                    Set Record2 = qViewCompID.Fetch()
                                                End If
                                                If Not Record2 Is Nothing Then 
                                                ' found the ComponentId
                                                ' this is a new file added to an existing component
                                                    arrFiles(FILES_COMPONENTID, iPosArr) = Record2.StringData(2)
                                                ' add the Directory_ reference
                                                    arrFiles(FILES_DIRECTORY, iPosArr) = Record2.StringData(3)
                                                    Set Record2 = Nothing
                                                    qViewCompID.Close
                                                    Set qViewCompID = Nothing
                                                Else
                                                ' did not find the ComponentId in the base .msi
                                                ' this is a new file AND a new component -> need to query the .msp for details
                                                    Set qViewMspCompId = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='Component' ORDER BY `Row`")
                                                    qViewMspCompID.Execute
                                                    Do
                                                        Set Record3 = qViewMspCompId.Fetch()
                                                        If Record3 Is Nothing Then Exit Do
                                                        If Record3.StringData(3) = arrFiles(FILES_COMPONENTNAME, iPosArr) Then 
                                                            If Record3.StringData(2) = "ComponentId" Then 
                                                                arrFiles(FILES_COMPONENTID, iPosArr) = Record3.StringData(4)
                                                            ElseIf Record3.StringData(2) = "Directory_" Then
                                                                arrFiles(FILES_DIRECTORY, iPosArr) = Record3.StringData(4)
                                                            End If
                                                        End If
                                                    Loop
                                                    qViewMspCompID.Close
                                                    If arrFiles(FILES_COMPONENTID, iPosArr) = "" Then bFtkForceOutOfScope = True
                                                End If 'Not Record2 Is Nothing
                                            ' all other logic is only needed if in scope
                                                If Not bFtkForceOutOfScope Then
                                                ' ensure the directory reference exists
                                                    If Not dicFolders.Exists(arrFiles(FILES_DIRECTORY, iPosArr)) Then
                                                        If SessionDir Is Nothing Then
                                                        ' try to recover lost SessionDir object
                                                            Set SessionDir = oMsi.OpenProduct(sProductCode)
                                                            SessionDir.DoAction("CostInitialize")
                                                            SessionDir.DoAction("FileCost")
                                                            SessionDir.DoAction("CostFinalize")
                                                        End If
                                                        dicFolders.Add arrFiles(FILES_DIRECTORY, iPosArr), SessionDir.TargetPath(arrFiles(FILES_DIRECTORY, iPosArr))
                                                        If Not Err = 0 Then 
                                                            Err.Clear
                                                        ' still failed to identify the path - get rid of this entry
                                                            bNeedKeyPathFallback = True 
                                                        End If
                                                    End If
                                                ' ComponentState
                                                    arrFiles(FILES_COMPONENTSTATE, iPosArr) = GetComponentState(sProductCode, arrFiles(FILES_COMPONENTID, iPosArr), iPosMaster)
                                                ' add ComponentClients
                                                    arrFiles(FILES_COMPONENTCLIENTS, iPosArr) = GetComponentClients(arrFiles(FILES_COMPONENTID, iPosArr), arrFiles(FILES_COMPONENTSTATE, iPosArr))
                                                ' add Features that use the component
                                                    Set qViewMspFC = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table`='FeatureComponents' ORDER BY `Row`")
                                                    qViewMspFC.Execute
                                                    Set Record2 = qViewMspFC.Fetch()
                                                    Do
                                                        Set Record2 = qViewMspFC.Fetch()
                                                        If Record2 Is Nothing Then Exit Do
                                                        If Record2.StringData(4) = arrFiles(FILES_COMPONENTNAME, iPosArr) Then 
                                                            If Record2.StringData(2) = "Feature_" Then 
                                                                arrFiles(FILES_FEATURENAMES, iPosArr) = arrFiles(FILES_FEATURENAMES, iPosArr) & Record2.StringData(3) &  _
                                                                "(" & TranslateFeatureState(oMsi.FeatureState(sProductCode, Record2.StringData(3))) & ")" &  ","
                                                            End If
                                                        End If
                                                    Loop
                                                    qViewMspFC.Close
                                                    Set Record2 = Nothing
                                                    RTrimComma arrFiles(FILES_FEATURENAMES, iPosArr)
                                                ' add KeyPath
                                                    arrFiles(FILES_KEYPATH, iPosArr) = GetComponentPath(sProductCode, arrFiles(FILES_COMPONENTID, iPosArr), arrFiles(FILES_COMPONENTSTATE, iPosArr))
                                                ' add Componentpath
                                                    sPath = "" : sName = ""
                                                    If dicAssembly.Exists(arrFiles(FILES_COMPONENTNAME, iPosArr)) Then
                                                    ' Assembly
                                                        If arrFiles(FILES_COMPONENTSTATE, iPosArr) = INSTALLSTATE_LOCAL Then
                                                            sPath = GetAssemblyPath(arrFiles(FILES_FILE, iPosArr), arrFiles(FILES_KEYPATH, iPosArr), dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr)))
                                                            arrFiles(FILES_FOLDER, iPosArr) = Left(sPath, InStrRev(sPath, "\"))
                                                            sName = Right(sPath, Len(sPath) -InStrRev(sPath, "\"))
                                                        ' update the files field to ensure the correct value
                                                            If Not UCase(sName) = UCase(arrFiles(FILES_FILE, iPosArr)) Then _
                                                                arrFiles(FILES_FILE, iPosArr) = sName 
                                                        Else
                                                            arrFiles(FILES_FOLDER, iPosArr) = sPath
                                                        End If
                                                    Else
                                                    ' Regular component
                                                        If bNeedKeyPathFallback Then
                                                            arrFiles(FILES_FOLDER, iPosArr) = Left(arrFiles(FILES_KEYPATH, iPosArr), InStrRev(arrFiles(FILES_KEYPATH, iPosArr), "\"))
                                                        Else
                                                            arrFiles(FILES_FOLDER, iPosArr) = dicFolders.Item(arrFiles(FILES_DIRECTORY, iPosArr))
                                                        End If
                                                    End If
                                                ' add file FullName
                                                ' if sPath contains a string then it's the result of the assembly detection
                                                    If sPath = "" Then
                                                        arrFiles(FILES_FULLNAME, iPosArr) = GetFileFullName(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FOLDER, iPosArr), arrFiles(FILES_FILE, iPosArr))
                                                    Else
                                                        arrFiles(FILES_FULLNAME, iPosArr) = arrFiles(FILES_FOLDER, iPosArr) & arrFiles(FILES_FILE, iPosArr)
                                                    End If
                                                ' add FileState and FileVersion
                                                    arrFiles(FILES_FILESTATUS, iPosArr) = GetFileState(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                                                    arrFiles(FILES_CURRENTVERSION, iPosArr) = GetFileVersion(arrFiles(FILES_COMPONENTSTATE, iPosArr), arrFiles(FILES_FULLNAME, iPosArr))
                                                End If
                                            End If 'bBaseRefFound
                                        Else
                                        ' No - FTK not in scope
                                            bFtkForceOutOfScope = False
                                        ' delete all row contents
                                            For iColCnt = 0 To FILES_COLUMNCNT
                                                arrFiles(iColCnt, iPosArr) = ""
                                            Next 'iColCnt
                                        ' decrease array counter
                                            iPosArr = iPosArr - 1
                                        End If 'bFtkInScope
                                        
                                        If bFtkForceOutOfScope Then
                                        ' delete all row contents
                                            For iColCnt = 0 To FILES_COLUMNCNT
                                                arrFiles(iColCnt, iPosArr) = ""
                                            Next 'iColCnt
                                        ' decrease array counter
                                            iPosArr = iPosArr - 1
                                        End If
                                    End If 'bFtkViewComplete
                                    bFtkViewComplete = True
                                    
                                ' Previous FTK handling is now complete
                                ' -------------------------------------
                                    If Record Is Nothing Then Exit Do
                                ' Init new FTK row
                                ' ----------------
                                ' increase array pointer
                                    iPosArr = iPosArr + 1
                                    bFtkViewComplete = False
                                    bFtkInScope = False
                                    bFtkForceOutOfScope = False
                                    bInsert = False
                                    bFileNameChanged = False
                                    If iPosArr > iArrMaxCnt Then
                                    ' add more rows to array
                                        iArrMaxCnt = iArrMaxCnt + 1000
                                        ReDim Preserve arrFiles(FILES_COLUMNCNT, iArrMaxCnt)
                                    End If 'iPosArr > iArrMaxCnt
                                ' update current FTK cache reference
                                    sFtk = Record.StringData(3)
                                ' locate the FTK reference from msi base
                                    bBaseRefFound = False
                                    For iCnt = 0 To iBaseRefCnt
                                        If arrFiles(FILES_FTK, iCnt) = sFtk Then
                                            bBaseRefFound = True
                                        ' copy known fields from Base version if applicable
                                            If bBaseRefFound Then
                                                For iColCnt = 0 To FILES_COLUMNCNT
                                                    arrFiles(iColCnt, iPosArr) = arrFiles(iColCnt, iCnt)
                                                Next 'iColCnt
                                                bFtkInScope = True
                                            End If 'bBaseRefFound
                                            Exit For 'iCnt = 0 To UBound(arrFiles, 2) - 1
                                        End If 'arrFiles(FILES_FTK, iCnt) = sFtk Then
                                    Next 'iCnt
                                ' add initial available data
                                ' correct/ensure FTK name
                                    arrFiles(FILES_FTK, iPosArr) = Record.StringData(3)
                                ' correct IsPatched field
                                    arrFiles(FILES_ISPATCHED, iPosArr) = False
                                ' correct/ensure FileSource
                                    arrFiles(FILES_SOURCE, iPosArr) = "Msp"
                                ' add fields from patch array
                                    arrFiles(FILES_PATCHSTATE, iPosArr) = arrPatch(iPosMaster, PATCH_PATCHSTATE, iPosPatch)
                                    arrFiles(FILES_PATCHCODE, iPosArr) = arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)
                                    arrFiles(FILES_PATCHMOREINFO, iPosArr) = arrPatch(iPosMaster, PATCH_MOREINFOURL, iPosPatch)
                                ' add KB reference
                                    If dicMspIndex.Exists(sPatchCode) Then
                                        iIndex = dicMspIndex.Item(sPatchCode)
                                        arrFiles(FILES_PATCHKB, iPosArr) = arrMspFiles(iIndex, MSPFILES_KB)
                                        arrFiles(FILES_PATCHPACKAGE, iPosArr) = arrMspFiles(iIndex, MSPFILES_PACKAGE)
                                    End If
                                ' new FTK row init complete
                                End If 'Not sFtk = Record.StringData(3)
                            ' add data from _TransformView
                                Select Case Record.StringData(2)
                                Case "File"
                                Case "FileSize"
                                Case "Component_"
                                    arrFiles(FILES_COMPONENTNAME, iPosArr) = Record.StringData(4)
                                Case "CREATE"
                                Case "DELETE"
                                Case "DROP"
                                Case "FileName"
                                'Add the filename
                                    bFileNameChanged = True
                                    arrFiles(FILES_FILE, iPosArr) = Record.StringData(4)
                                Case "Version"
                                ' don't allow version field to contain alpha characters
                                    bAsc = True : sAscCheck = "" : sAscCheck = Record.StringData(4)
                                    If Len(sAscCheck) > 0 Then
                                        For iAscCnt = 1 To Len(sAscCheck)
                                            iAsc = Asc(UCase(Mid(sAscCheck, iAscCnt, 1)))
                                            If (iAsc>64) AND (iAsc<91) Then
                                                bAsc = False
                                                Exit For
                                            End If
                                        Next 'iCnt
                                    End If 'Len(sAscCheck) > 0
                                    If bAsc Then arrFiles(FILES_PATCHVERSION, iPosArr) = Record.StringData(4)
                                Case "Language"
                                        arrFiles(FILES_LANGUAGE, iPosArr) = Record.StringData(4)
                                Case "Attributes"
                                Case "Sequence"
                                Case "INSERT"
                                ' this is a new file added by the pach
                                    bFtkInScope = True
                                    bInsert = True
                                Case Else
                                End Select
                            ' get the next record (column) from _TransformView
                                Set Record = qViewMst.Fetch()
                            Loop
                        ' _TransformView analysis for this patch complete
                        ' reset views 
                            MspDb.OpenView("ALTER TABLE _TransformView FREE").Execute
                            MspDb.OpenView("DROP TABLE `File`").Execute
                        End If 'Err = 0
                    End If 'IsEmpty
                Next 'iPosPatch
            Next 'iFoo
            
        ' Final field & verb fixups
        ' -------------------------
            For iFile = 0 To iPosArr
            ' File VersionState translation
                iCmp = 2 : sCmp = "" : sCurVer = ""
                If arrFiles(FILES_ISPATCHED, iFile) OR (arrFiles(FILES_SOURCE, iFile) = "Msp") Then 
                ' compare actual file version to patch file version 
                    sCurVer = arrFiles(FILES_PATCHVERSION, iFile)
                Else
                ' compare actual file version to base file version 
                    sCurVer = arrFiles(FILES_BASEVERSION, iFile)
                End If 'arrFiles(FILES_ISPATCHED, iFile)
                iCmp = CompareVersion(arrFiles(FILES_CURRENTVERSION, iFile), sCurVer, False)
                Select Case iCmp
                Case VERSIONCOMPARE_LOWER
                    sCmp= "ERROR_VersionLow"
                ' log error
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_FILEVERSIONLOW & arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & _
                    arrMaster(iPosMaster, COL_PRODUCTNAME)
                    If Not InStr(arrMaster(iPosMaster, COL_ERROR), arrFiles(FILES_FILE, iFile) & " expected: ") > 0 Then
                        arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & arrFiles(FILES_FILE, iFile) & " expected: " & sCurVer &  " found: " & arrFiles(FILES_CURRENTVERSION, iFile) & CSV
                    End If
                Case VERSIONCOMPARE_MATCH   : sCmp= "SUCCESS_VersionMatch"
                Case VERSIONCOMPARE_HIGHER  : sCmp= "SUCCESS_VersionHigh"
                Case VERSIONCOMPARE_INVALID : sCmp= ""
                End Select
                arrFiles(FILES_VERSIONSTATUS, iFile) = sCmp
            ' FileState translation
                sCmp = ""
                Select Case arrFiles(FILES_FILESTATUS, iFile)
                Case INSTALLSTATE_LOCAL : sCmp = "OK_Local"
                Case INSTALLSTATE_BROKEN
                    sCmp = "ERROR_Broken"
                ' log error
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & _
                    arrMaster(iPosMaster, COL_PRODUCTNAME) & ": " & ERR_BADFILESTATE
                    If Not InStr(arrMaster(iPosMaster, COL_ERROR), arrFiles(FILES_FILE, iFile) & " FileState: Broken") > 0 Then
                        arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & arrFiles(FILES_FILE, iFile) & " FileState: Broken" & CSV
                    End If
                Case INSTALLSTATE_UNKNOWN
                    sCmp = "Unknown"
                    If Not arrFiles(FILES_FEATURENAMES, iFile) = "" Then sCmp=Mid(arrFiles(FILES_FEATURENAMES, iFile), InStrRev(arrFiles(FILES_FEATURENAMES, iFile), "(") + 1, Len(arrFiles(FILES_FEATURENAMES, iFile)) -InStrRev(arrFiles(FILES_FEATURENAMES, iFile), "(") - 1)
                Case INSTALLSTATE_NOTUSED : sCmp = "NotUsed"
                Case INSTALLSTATE_ASSEMBLY : sCmp = "Assembly"
                Case Else
                End Select
                arrFiles(FILES_FILESTATUS, iFile) = sCmp
            ' ComponentState translation
                sCmp = ""
                Select Case arrFiles(FILES_COMPONENTSTATE, iFile)
                Case INSTALLSTATE_LOCAL : sCmp = "Local"
                Case INSTALLSTATE_BROKEN : sCmp = "Broken"
                Case INSTALLSTATE_UNKNOWN : sCmp = "Unknown"
                Case INSTALLSTATE_NOTUSED : sCmp = "NotUsed"
                Case Else
                End Select
                arrFiles(FILES_COMPONENTSTATE, iFile) = sCmp
            ' PatchCode field trim
                arrFiles(FILES_PATCHCODE, iFile) = RTrimComma(arrFiles(FILES_PATCHCODE, iFile))
            ' PatchKB field trim
                arrFiles(FILES_PATCHKB, iFile) = RTrimComma(arrFiles(FILES_PATCHKB, iFile))
            ' PatchInfo field trim
                arrFiles(FILES_PATCHMOREINFO, iFile) = RTrimComma(arrFiles(FILES_PATCHMOREINFO, iFile))
            Next 'iFile
        ' dump out the collected data to file
        ' create the AllOffice file
            If IsEmpty(AllOfficeFiles) Then
                If NOT oFso.FolderExists(sPathOutputFolder & "ROIScan") Then oFso.CreateFolder(sPathOutputFolder & "ROIScan")
                Set AllOfficeFiles = oFso.CreateTextFile(sPathOutputFolder & "ROIScan\" & sComputerName & "_OfficeAll_FileList.xml", True, True)
                AllOfficeFiles.WriteLine "<?xml version= ""1.0""?>"
                AllOfficeFiles.WriteLine "<FILEDATA>"
            End If
        ' individual products file
            If NOT oFso.FolderExists(sPathOutputFolder & "ROIScan") Then oFso.CreateFolder(sPathOutputFolder & "ROIScan")
            Set FileStream = oFso.CreateTextFile(sPathOutputFolder & "ROIScan\" & sComputerName & "_" & sProductCode & "_FileList.xml", True, True)
            FileStream.WriteLine "<?xml version= ""1.0""?>"
            FileStream.WriteLine "<FILEDATA>"
            FileStream.WriteLine vbTab & "<PRODUCT ProductCode=""" & sProductCode & """ >"
            If arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine vbTab & "<PRODUCT ProductCode=""" & sProductCode & """ >"
            For iFile = 0 To iPosArr
                sXmlLine = ""
                sXmlLine = vbTab & vbTab & "<FILE " & _
                                           "FileName=" & chr(34) & arrFiles(FILES_FILE, iFile) & chr(34) & " " & _
                                           "FileState=" & chr(34) & arrFiles(FILES_FILESTATUS, iFile) & chr(34) & " " & _
                                           "VersionStatus=" & chr(34) & arrFiles(FILES_VERSIONSTATUS, iFile) & chr(34) & " " & _
                                           "CurrentVersion=" & chr(34) & arrFiles(FILES_CURRENTVERSION, iFile) & chr(34) & " " & _
                                           "InitialVersion=" & chr(34) & arrFiles(FILES_BASEVERSION, iFile) & chr(34) & " " & _
                                           "PatchVersion=" & chr(34) & arrFiles(FILES_PATCHVERSION, iFile) & chr(34) & " " & _
                                           "FileSource=" & chr(34) & arrFiles(FILES_SOURCE, iFile) & chr(34) & " " & _
                                           "IsPatched=" & chr(34) & arrFiles(FILES_ISPATCHED, iFile) & chr(34) & " " & _
                                           "KB=" & chr(34) & arrFiles(FILES_PATCHKB, iFile) & chr(34) & " " & _
                                           "Package=" & chr(34) & arrFiles(FILES_PATCHPACKAGE, iFile) & chr(34) & " " & _
                                           "PatchState=" & chr(34) & arrFiles(FILES_PATCHSTATE, iFile) & chr(34) & " " & _
                                           "FolderName=" & chr(34) & arrFiles(FILES_FOLDER, iFile) & chr(34) & " " & _
                                           "PatchCode=" & chr(34) & arrFiles(FILES_PATCHCODE, iFile) & chr(34) & " " & _
                                           "PatchInfo=" & chr(34) & arrFiles(FILES_PATCHMOREINFO, iFile) & chr(34) & " " & _
                                           "FtkName=" & chr(34) & arrFiles(FILES_FTK, iFile) & chr(34) & " " & _
                                           "KeyPath=" & chr(34) & arrFiles(FILES_KEYPATH, iFile) & chr(34) & " " & _
                                           "MsiDirectory=" & chr(34) & arrFiles(FILES_DIRECTORY, iFile) & chr(34) & " " & _
                                           "Language=" & chr(34) & arrFiles(FILES_LANGUAGE, iFile) & chr(34) & " " & _
                                           "ComponentState=" & chr(34) & arrFiles(FILES_COMPONENTSTATE, iFile) & chr(34) & " " & _
                                           "ComponentID=" & chr(34) & arrFiles(FILES_COMPONENTID, iFile) & chr(34) & " " & _
                                           "ComponentName=" & chr(34) & arrFiles(FILES_COMPONENTNAME, iFile) & chr(34) & " " & _
                                           "ComponentClients=" & chr(34) & arrFiles(FILES_COMPONENTCLIENTS, iFile) & chr(34) & " " & _
                                           "FeatureReference=" & chr(34) & arrFiles(FILES_FEATURENAMES, iFile) & chr(34) & " " & _
                                           " />"
                If InStr(sXmlLine, "& ") > 0 Then sXmlLine = Replace(sXmlLine, "& ","&amp;")
                FileStream.WriteLine sXmlLine
                If arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine sXmlLine
            Next 'iFile
            FileStream.WriteLine vbTab & "</PRODUCT>"
            If arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) Then AllOfficeFiles.WriteLine vbTab & "</PRODUCT>"
            FileStream.WriteLine "</FILEDATA>"
            FileStream.Close
            Set FileStream = Nothing
        End If 'arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)
    Next 'iPosMaster
' close the AllOffice file
    If Not AllOfficeFiles Is Nothing Then
        AllOfficeFiles.WriteLine "</FILEDATA>"
        AllOfficeFiles.Close
        Set AllOfficeFiles = Nothing
    End if
' compress the files 
    Dim i, iWait
    Dim FileScanFolder, FileVerScanZip, xmlFile, item, zipfile
    Dim sDat, sDatCln
    Dim fCopyComplete
    If oFso.FileExists(sPathOutputFolder & sComputerName & "_ROIScan.zip") Then
    ' rename existing .zip container by appending a timestamp to prevent overwrite.
        Dim oRegExp
        Set oRegExp = CreateObject("Vbscript.RegExp")
        Set zipfile = oFso.GetFile(sPathOutputFolder & sComputerName & "_ROIScan.zip")
        oRegExp.Global = True
        oRegExp.Pattern = "\D"
        Err.Clear
        zipfile.Name = sComputerName & "_ROIScan_" & oRegExp.Replace(zipfile.DateLastModified, "") & ".zip"
        If NOT Err = 0 Then
            zipfile.Delete
            Err.Clear
        End If
    End If
    Set FileVerScanZip = oFso.OpenTextFile(sPathOutputFolder & sComputerName & "_ROIScan.zip", FOR_WRITING, True) 
    FileVerScanZip.write "PK" & chr(5) & chr(6) & String(18, chr(0)) 
    FileVerScanZip.close  
    Set FileScanFolder = oFso.GetFolder(sPathOutputFolder & "ROIScan")
    For Each xmlFile in FileScanFolder.Files 
        If Right(LCase(xmlFile.Name), 4) = ".xml" Then
            If NOT fZipError Then CopyToZip ShellApp.NameSpace(sPathOutputFolder & sComputerName & "_ROIScan.zip"), xmlFile
        End If 
    Next 'xmlFile 
    If fCScript AND NOT fQuiet Then wscript.echo vbTab & "File version scan complete"

End Sub 'FileInventory
'=======================================================================================================

'Identify the InstallState of a component
Function GetComponentState(sProductCode, sComponentId, iPosMaster)
    On Error Resume Next
    
    Dim Product
    Dim sPath

    GetComponentState = INSTALLSTATE_UNKNOWN
    If iWiVersionMajor > 2 Then
        'WI 3.x or higher
        Set Product = oMsi.Product(sProductCode, arrMaster(iPosMaster, COL_USERSID), arrMaster(iPosMaster, COL_CONTEXT))
        Err.Clear
        GetComponentState = Product.ComponentState(sComponentId)
        If Not Err = 0 Then 
            GetComponentState = INSTALLSTATE_UNKNOWN
            Err.Clear
        End If ' Err = 0
    Else
        'WI 2.x
        If Not Err = 0 Then Err.Clear
        sPath = ""
        sPath = oMsi.ComponentPath(sProductCode, sComponentId)
        If Not Err = 0 Then
            GetComponentState = INSTALLSTATE_UNKNOWN
            Err.Clear
        Else
            If oFso.FileExists(sPath) Then
                GetComponentState = INSTALLSTATE_LOCAL
            Else
                GetComponentState = INSTALLSTATE_NOTUSED
            End If 'oFso.FileExists(sPath)
        End If 'Not Err = 0
    End If 'iWiVersionMajor > 2

End Function 'GetComponentState
'=======================================================================================================

'Get a list of client products that are registered to the component
Function GetComponentClients(sComponentId, iComponentState)
    On Error Resume Next
    
    Dim sClients, prod

    sClients = ""
    GetComponentClients = ""
    If Not(iComponentState = INSTALLSTATE_UNKNOWN) Then
        For Each prod in oMsi.ComponentClients(sComponentId)
            If Not Err = 0 Then 
                Err.Clear
                Exit For
            End If 'Not Err = 0
            sClients = sClients & prod & ","
        Next 'prod
        RTrimComma sClients
        GetComponentClients = sClients
    End If 'Not (arrFiles(FILES_COMPONENTSTATE, ...

End Function 'GetComponentClients
'=======================================================================================================

'Get the keypath value for the component
Function GetComponentPath(sProductCode, sComponentId, iComponentState)
    On Error Resume Next
    
    Dim sPath

    sPath = "" 
    If iComponentState = INSTALLSTATE_LOCAL Then
        sPath = oMsi.ComponentPath(sProductCode, sComponentId)
    End If 'iComponentState = INSTALLSTATE_LOCAL
    GetComponentPath = sPath

End Function 'GetComponentPath
'=======================================================================================================


'Use WI ProvideAssembly function to identify the path for an assembly.
'Returns the path to the file if the file exists.
'Returns an empty string if file does not exist

Function GetAssemblyPath(sLfn, sKeyPath, sDir)
    On Error Resume Next
    
    Dim sFile, sFolder, sExt, sRoot, sName
    Dim arrTmp
    
    'Defaults
    GetAssemblyPath= ""
    sFile="" : sFolder= "" : sExt= "" : sRoot= "" : sName=""
    

    'The componentpath should already point to the correct folder
    'except for components with a registry keypath element.
    'In that case tweak the directory folder to match
    If Left(sKeyPath, 1) = "0" Then
        sFolder = sDir
        sFolder = oShell.ExpandEnvironmentStrings("%SYSTEMROOT%") & Mid(sFolder, InStr(LCase(sFolder), "\winsxs\"))
        sFile = sLfn
    End If 'Left(sKeyPath, 1) = "0"
    
    'Figure out the correct file reference
    If sFolder = "" Then sFolder = Left(sKeyPath, InStrRev(sKeyPath, "\"))
    sRoot = Left(sFolder, InStrRev(sFolder, "\", Len(sFolder) - 1))
    arrTmp = Split(sFolder, "\")
    If CheckArray(arrTmp) Then sName = arrTmp(UBound(arrTmp) - 1)
    If sFile = "" Then sFile = Right(sKeyPath, Len(sKeyPath) -InStrRev(sKeyPath, "\"))
    If oFso.FileExists(sFolder & sLfn) Then 
        sFile = sLfn
    Else
        'Handle .cat, .manifest and .policy files
        If InStr(sLfn, ".") > 0 Then
            sExt = Mid(sLfn, InStrRev(sLfn, "."))
            Select Case LCase(sExt)
            Case ".cat"
                sFile = Left(sFile, InStrRev(sFile, ".")) & "cat"
                If Not oFso.FileExists(sFolder & sFile) Then
                    'Check Manifest folder
                    If oFso.FileExists(sRoot & "Manifests\" & sName & ".cat") Then
                        sFolder = sRoot & "Manifests\"
                        sFile = sName & ".cat"
                    Else
                        If oFso.FileExists(sRoot & "Policies\" & sName & ".cat") Then
                            sFolder = sRoot & "Policies\"
                            sFile = sName & ".cat"
                        End If
                    End If
                End If
            Case ".manifest"
                sFile = Left(sFile, InStrRev(sFile, ".")) & "manifest"
                If oFso.FileExists(sRoot & "Manifests\" & sName & ".manifest") Then
                    sFolder = sRoot & "Manifests\"
                    sFile = sName & ".manifest"
                End If
            Case ".policy"
                If iVersionNT < 600 Then
                    sFile = Left(sFile, InStrRev(sFile, ".")) & "policy"
                    If oFso.FileExists(sRoot & "Policies\" & sName & ".policy") Then
                        sFolder = sRoot & "Policies\"
                        sFile = sName & ".policy"
                    End If
                Else
                    sFile = Left(sFile, InStrRev(sFile, ".")) & "manifest"
                    If oFso.FileExists(sRoot & "Manifests\" & sName & ".manifest") Then
                        sFolder = sRoot & "Manifests\"
                        sFile = sName & ".manifest"
                    End If
                End If
            Case Else
            End Select
            
            'Check if the file exists
            If Not oFso.FileExists(sFolder & sFile) Then
                'Ensure the right folder
            End If
        End If 'InStr(sFile, ".") > 0
    End If
    
    GetAssemblyPath = sFolder & sFile

End Function 'GetAssemblyPath
'=======================================================================================================


Function GetFileFullName(iComponentState, sComponentPath, sFileName)
    On Error Resume Next
    
    Dim sFileFullName

    sFileFullName = ""
    If iComponentState = INSTALLSTATE_LOCAL Then
        If Len(sComponentPath) > 2 Then sFileFullName = sComponentPath & sFileName
    End If 'iComponentState = INSTALLSTATE_LOCAL
    GetFileFullName = sFileFullName

End Function 'GetFileFullName
'=======================================================================================================

Function GetLongFileName(sMsiFileName)
    On Error Resume Next
    
    Dim sFileTmp
    
    sFileTmp = ""
    sFileTmp = sMsiFileName
    If InStr(sFileTmp, "|") > 0 Then sFileTmp = Mid(sFileTmp, InStr(sFileTmp, "|") + 1, Len(sFileTmp))
    GetLongFileName = sFileTmp

End Function 'GetLongFileName
'=======================================================================================================

Function GetFileState(iComponentState, sFileFullName)
    On Error Resume Next
    
    GetFileState = INSTALLSTATE_UNKNOWN
    If iComponentState = INSTALLSTATE_LOCAL Then
        If oFso.FileExists(sFileFullName) Then
            GetFileState = INSTALLSTATE_LOCAL
        Else
            GetFileState = INSTALLSTATE_BROKEN
        End If 'oFso.FileExists(sFileFullName) 
    Else
        If oFso.FileExists(sFileFullName) Then
            'This should not happen!
            GetFileState = INSTALLSTATE_LOCAL
        Else
            GetFileState = iComponentState
        End If 'oFso.FileExists(sFileFullName) 
    End If 'iComponentState = INSTALLSTATE_LOCAL

End Function 'GetFileState
'=======================================================================================================

Function GetFileVersion(iComponentState, sFileFullName)
    On Error Resume Next
    
    GetFileVersion = ""
    
    If iComponentState = INSTALLSTATE_LOCAL Then
        If oFso.FileExists(sFileFullName) Then
            GetFileVersion = oFso.GetFileVersion(sFileFullName)
        End If 'oFso.FileExists(sFileFullName) 
    Else
        If oFso.FileExists(sFileFullName) Then
            'This should not happen!
            GetFileVersion = oFso.GetFileVersion(sFileFullName)
        End If 'oFso.FileExists(sFileFullName) 
    End If 'iComponentState = INSTALLSTATE_LOCAL

End Function 'GetFileVersion
'=======================================================================================================


'=======================================================================================================
'Module FeatureStates
'=======================================================================================================

'Builds a FeatureTree indicating the FeatureStates 
Sub FindFeatureStates
    If fBasicMode Then Exit Sub
    On Error Resume Next

    Const ADVARCHAR     = 200
    Const MAXCHARACTERS = 255

    Dim Features, oRecordSet, oDicLevel, oDicParent
    Dim sProductCode, sFeature, sFTree, sFParent, sLeft, sRight
    Dim iFoo, iPosMaster, iMaxNestLevel, iNestLevel, iLevel, iFCnt, iLeft, iStart
    Dim arrFName, arrFLevel, arrFParent
    
    'ReDim the global feature array
    ReDim arrFeature (UBound (arrMaster), FEATURE_COLUMNCOUNT)
    
    'Outer loop to iterate through all products
    For iPosMaster = 0 To UBound (arrFeature)
        iFoo = 0
        sFTree = ""
        'Dummy Loop to allow exit out in case of an error
        Do While iFoo = 0
            iFoo = 1
            'Get the ProductCode for this loop
            sProductCode = arrMaster (iPosMaster, COL_PRODUCTCODE)
            'Get the Features colection for this product
            
            'oMsi.Features is only valid for installed per-machine or current user products.
            'The call will fail for advertised and other user products.
            Set Features = oMsi.Features (sProductCode)
            If Not Err = 0 Then
                Err.Clear
                Exit Do
            End If 'Not Err = 0
            'Create the dictionary objects
            Set oDicLevel  = CreateObject ("Scripting.Dictionary")
            Set oDicParent = CreateObject ("Scripting.Dictionary")
            'Prepare a recordset to allow sorting of the root features
            Set oRecordSet = CreateObject ("ADOR.Recordset")
            oRecordSet.Fields.Append "FeatureName", ADVARCHAR, MAXCHARACTERS
            oRecordSet.Open
            If Not Err = 0 Then
                Err.Clear
                Exit Do
            End If 'Not Err = 0
            iMaxNestLevel = 0
            
            'Inner loop # 1 to identify all features
            For Each sFeature in Features
                'Reset the nested level counter
                iNestLevel = -1
                'Check for & cache parent feature
                sFParent = oMsi.FeatureParent (sProductCode, sFeature)
                If sFParent = "" Then 
                    'Found a root feature.
                    iNestLevel = 0
                    'Add to recordset for later sorting
                    oRecordSet.AddNew
                    oRecordSet("FeatureName") = TEXTINDENT & sFeature & " (" & TranslateFeatureState (oMsi.FeatureState (sProductCode, sFeature)) & ")"
                    oRecordSet.Update
                Else
                    'Call the recursive function to get the nest level of the current feature
                    iNestLevel = GetFeatureNestLevel (sProductCode, sFeature, iNestLevel)
                    'Add to dictionary arrays
                    oDicLevel.Add sFeature, iNestLevel
                    oDicParent.Add sFeature, sFParent
                End If 'sFParent= ""
                'Max nest level is required for second inner loop
                If iNestLevel > iMaxNestLevel Then iMaxNestLevel = iNestLevel
            Next 'sFeature
            
            'First inner loop complete. Sort the root features
            oRecordSet.Sort = "FeatureName"
            oRecordSet.MoveFirst
            'Write the sorted root features to the 'treeview' string
            Do Until oRecordSet.EOF
                sFTree = sFTree & oRecordSet.Fields.Item ("FeatureName") & vbCrLf
                oRecordSet.MoveNext
            Loop 'oRecordSet.EOF
            
            'Copy dic's to array
            arrFName  = oDicLevel.Keys
            arrFLevel = oDicLevel.Items
            arrFParent= oDicParent.Items
            
            '2nd inner loop to add the features to the 'treeview' string
            For iLevel = 1 To iMaxNestLevel
                For iFCnt = 0 To UBound(arrFName)
                    If arrFLevel (iFCnt) = iLevel Then
                        iStart = InStr (sFTree, arrFParent (iFCnt) & " (") + Len (arrFParent (iFCnt))
                        iLeft  = InStr (iStart, sFTree, ")") + 2
                        sLeft  = Left (sFTree, iLeft)
                        sRight = Right (sFTree, Len (sFTree) - iLeft)
                        sFTree = sLeft & TEXTINDENT & FeatureIndent (iLevel) & arrFName (iFCnt) & " (" & TranslateFeatureState (oMsi.FeatureState (sProductCode, arrFName (iFCnt))) & ")" & vbCrLf & sRight
                    End If 'arrFLevel(iFCnt) =i
                Next 'iFCnt
            Next 'iLevel
            
            'Reset objects for next cycle
            Set oRecordSet = Nothing
            Set oDicLevel  = Nothing
            Set oDicParent = Nothing
        Loop 'iFoo= 0
        
        arrFeature (iPosMaster, FEATURE_TREE) = vbCrLf & sFTree
    Next 'iProdMaster

End Sub 'FindFeatureStates
'=======================================================================================================

'Translate the FeatureState value
Function TranslateFeatureState(iFState)

    Select Case iFState
    Case INSTALLSTATE_UNKNOWN       : TranslateFeatureState="Unknown"
    Case INSTALLSTATE_ADVERTISED    : TranslateFeatureState="Advertised"
    Case INSTALLSTATE_ABSENT        : TranslateFeatureState="Absent"
    Case INSTALLSTATE_LOCAL         : TranslateFeatureState="Local"
    Case INSTALLSTATE_SOURCE        : TranslateFeatureState="Source"
    Case INSTALLSTATE_DEFAULT       : TranslateFeatureState="Default"
    Case INSTALLSTATE_VIRTUALIZED   : TranslateFeatureState="Virtualized"
    Case INSTALLSTATE_BADCONFIG     : TranslateFeatureState="BadConfig"
    Case Else                       : TranslateFeatureState="Error"
    End Select

End Function 'GetFeatureStateString
'=======================================================================================================
Function FeatureIndent(iNestLevel)
    Dim iLevel
    Dim sIndent

    For iLevel = 1 To iNestLevel
        sIndent = sIndent & vbTab
    Next 'iLevel
    FeatureIndent = sIndent
End Function 'FeatureIndent
'=======================================================================================================

Function GetFeatureNestLevel(sProductCode, sFeature, iNestLevel)
    Dim sParent : sParent = ""
    iNestLevel=iNestLevel+1
    sParent=oMsi.FeatureParent(sProductCode, sFeature)
    If Not sParent = "" Then iNestLevel=GetFeatureNestLevel(sProductCode, sParent, iNestLevel)
    GetFeatureNestLevel = iNestLevel
End Function 'GetFeatureNestLevel
'=======================================================================================================


'=======================================================================================================
'Module Product InstallSource - 
'=======================================================================================================

Sub ReadMsiInstallSources ()
    If fBasicMode Then Exit Sub
    On Error Resume Next


    Dim oProduct, oSumInfo
    Dim iProdCnt, iSourceCnt
    Dim sSource
    Dim MsiSources

    ReDim arrIS(UBound(arrMaster), UBOUND_IS)

    For iProdCnt = 0 To UBound(arrMaster)
        arrIS(iProdCnt, IS_SOURCETYPESTRING) = "No Data Available"
        arrIS(iProdCnt, IS_ORIGINALSOURCE)   = "No Data Available"
        'Add the ProductCode to the array
        arrIS(iProdCnt, IS_PRODUCTCODE) = arrMaster(iProdCnt, COL_PRODUCTCODE)
        
        If arrMaster(iProdCnt, COL_VIRTUALIZED) = 1 Then
            ' do nothing
        Else
            'SourceType
            If oFso.FileExists(arrMaster(iProdCnt, COL_CACHEDMSI)) Then
                Err.Clear
                Set oSumInfo = oMsi.SummaryInformation(arrMaster(iProdCnt, COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
                If Err = 0 Then
                    arrIS(iProdCnt, IS_SOURCETYPE) = oSumInfo.Property(PID_WORDCOUNT)
                    Select Case arrIS(iProdCnt, IS_SOURCETYPE)
                    Case 0 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Original source using long file names"
                    Case 1 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Original source using short file names"
                    Case 2 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Compressed source files using long file names"
                    Case 3 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Compressed source files using short file names"
                    Case 4 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Administrative image using long file names"
                    Case 5 : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Administrative image using short file names"
                    Case Else : arrIS(iProdCnt, IS_SOURCETYPESTRING) = "Unknown InstallSource Type"
                    End Select
                Else
                    'ERR_SICONNECTFAILED
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & arrMaster(iProdCnt, COL_PRODUCTCODE) & DSV & _
                    arrMaster(iProdCnt, COL_PRODUCTNAME) & ": " & ERR_SICONNECTFAILED
                    arrMaster(iProdCnt, COL_ERROR) = arrMaster(iProdCnt, COL_ERROR) & ERR_CATEGORYERROR & ERR_SICONNECTFAILED & CSV
                End If 'Err
            End If
        
            'Get the original InstallSource
            arrIS(iProdCnt, IS_ORIGINALSOURCE) = oMsi.ProductInfo(arrMaster(iProdCnt, COL_PRODUCTCODE), "InstallSource")
            If Not Len(arrIS(iProdCnt, IS_ORIGINALSOURCE)) > 0 Then arrIS(iProdCnt, IS_ORIGINALSOURCE) = "Not Registered"
        
            'Get Network InstallSource(s)
            'With WI 3.x and later the 'Product' object can be used to gather some data
            If iWiVersionMajor > 2 Then
                Err.Clear
                Set oProduct = oMsi.Product(arrMaster(iProdCnt, COL_PRODUCTCODE), arrMaster(iProdCnt, COL_USERSID), arrMaster(iProdCnt, COL_CONTEXT))
                If Err = 0 Then
                    'Get the last used source
                    arrIS(iProdCnt, IS_LASTUSEDSOURCE) = oProduct.SourceListInfo("LastUsedSource")
            
                    Set MsiSources = oProduct.Sources(1)
                    For Each sSource in MsiSources
                        If IsEmpty(arrIS(iProdCnt, IS_ADDITIONALSOURCES)) Then
                            arrIS(iProdCnt, IS_ADDITIONALSOURCES) =sSource 'MsiSources(iSourceCnt)
                        Else
                            arrIS(iProdCnt, IS_ADDITIONALSOURCES) =arrIS(iProdCnt, IS_ADDITIONALSOURCES) & " || " &  sSource'MsiSources(iSourceCnt)
                        End If
                    Next 'MsiSources
                End If 'Err
            End If 'iWiVersionMajor
        
            'Get the LIS resiliency source (if applicable)
            If GetDeliveryResiliencySource(arrMaster(iProdCnt, COL_PRODUCTCODE), iProdCnt, sSource) Then arrIS(iProdCnt, IS_LISRESILIENCY) =sSource
        End If 'Not Virtualized        
    Next 'iProdCnt
End Sub 'ReadMsiInstallSources

'=======================================================================================================

'Return True/False and the LIS source path as sSource
'Empty string for sProductCode forces to identify the DownloadCode from Setup.xml
Function GetDeliveryResiliencySource (sProductCode, iPosMaster, sSource)
    On Error Resume Next
    
    Dim arrSources, arrDownloadCodeKeys
    Dim dicDownloadCode
    Dim sSubKeyName, sValue, key, sku, sSkuName, sText, sDownloadCode, sTmpDownloadCode, source
    Dim arrKeys, arrSku
    Dim iVersionMajor, iSrc
    Dim fFound

    GetDeliveryResiliencySource = False
    sSource = Empty
    iVersionMajor = GetVersionMajor(sProductCode)
    Set dicDownloadCode = CreateObject("Scripting.Dictionary")
    
    If iVersionMajor > 11 Then
        'Note: ProductCode doesn't work consistently for this logic
        '      To locate the Setup.xml requires additional logic so the tweak here is to use the 
        '      original source location to identify the DownloadCode
        sText = arrIS(iPosMaster, IS_ORIGINALSOURCE)
        If InStr(source, "{") > 0 Then sDownloadCode = Mid(sText, InStr(sText, "{"), 40) Else sDownloadCode = sProductCode
        dicDownloadCode.Add sDownloadCode, sProductCode
        
        'Find the additional download locations
        'Check if more than one sources are registered
        If InStr(arrIS(iPosMaster, IS_ADDITIONALSOURCES), "||") > 0 Then
            arrSources = Split(arrIS(iPosMaster, IS_ADDITIONALSOURCES), " || ")
            For Each source in arrSources
                If InStr(source, "{") > 0 Then
                    sTmpDownloadCode = Mid(source, InStr(source, "{"), 40)
                    If Not dicDownloadCode.Exists(sTmpDownloadCode) Then dicDownloadCode.Add sTmpDownloadCode, sProductCode
                End If 'InStr
            Next'
        End If 'InStr
        
        
        arrDownloadCodeKeys = dicDownloadCode.Keys
        For iSrc = 0 To dicDownloadCode.Count- 1
            sDownloadCode = UCase(arrDownloadCodeKeys(iSrc))
            'Enum HKLM\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads
            sSubKeyName="SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\"
            If RegEnumKey(HKLM, sSubKeyName, arrKeys) Then
                For Each key in arrKeys
                    fFound = False
                    If Len(key) > 37 Then
                        fFound = (UCase(Left(key, 38)) = sDownloadCode) OR (UCase(key) = sDownloadCode)
                    Else
                        fFound = (UCase(key) = sDownloadCode)
                    End If 'Len > 37
                    If fFound Then
                        'Found the Delivery reference
                        'Enum the 'Sources' subkey
                        sSubKeyName = sSubKeyName & key & "\Sources\"
                        If RegEnumKey(HKLM, sSubKeyName, arrSku) Then
                            For Each sku in arrSku
                                If RegReadStringValue(HKLM, sSubKeyName & sku, "Path", sValue) Then
                                    sSkuName = ""
                                    sSkuName = " (" & Left(sku, InStr(sku, "(") - 1) & ")"
                                    If IsEmpty(sSource) Then
                                        sSource = sValue & sSkuName
                                    Else
                                        sSource = sSource & " || " & sValue & sSkuName
                                    End If 'IsEmpty
                                End If 'RegReadStringValue
                            Next 'sku
                        End If 'RegEnumKey
                        
                        'GUID is unique no need to continue loop once we found a match
                        Exit For
                    End If
                Next 'key
            End If 'RegEnumKey
        Next 'iSrc
        
    ElseIf iVersionMajor = 11 Then
        'Get the DownloadCode
        sSubKeyName = "SOFTWARE\Microsoft\Office\11.0\Delivery\" & sProductCode & "\"
        If RegReadStringValue(HKLM, sSubKeyName, "DownloadCode", sDownloadCode) Then
            sSubKeyName = "SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\" & sDownloadCode & "\Sources\" & Mid(sProductCode, 2, 36) & "\"
            If RegReadStringValue(HKLM, sSubKeyName, "Path", sValue) Then sSource = sValue
        End If 
    End If 'iVersionMajor
    
    If Not IsEmpty(sSource) Then
        GetDeliveryResiliencySource = True
        WriteDebug sActiveSub, "Delivery resiliency source for " & sProductCode & " returned 'TRUE': " & sSource
    Else
        WriteDebug sActiveSub, "Delivery resiliency source for " & sProductCode & " returned 'FALSE'"
    End If

End Function 'GetDeliveryResiliencySource
'=======================================================================================================

'=======================================================================================================
'Module Product Properties
'=======================================================================================================

'Gather additional properties for the product
Sub ProductProperties
    Dim prod, MsiDb
    Dim iPosMaster, iVersionMajor, iContext, iProd
    Dim sProductCode, sSpLevel, sCachedMsi, sSid, sComp, sComp2, sPath, sRef, sProdId, n
    Dim fVer, fCx
    Dim arrCx, arrCxN, arrCxErr, arrKeys, arrTypes, arrNames
    On Error Resume Next

    If NOT fInitArrProdVer Then InitProdVerArrays 
    For iPosMaster = 0 to UBound (arrMaster)
    ' collect properties only for products in state '5' (Default)
        If arrMaster(iPosMaster, COL_STATE) = INSTALLSTATE_DEFAULT OR arrMaster(iPosMaster, COL_STATE) = INSTALLSTATE_VIRTUALIZED Then
            sProductCode = arrMaster(iPosMaster, COL_PRODUCTCODE)
            sSid = arrMaster(iPosMaster, COL_USERSID)
            iContext = arrMaster(iPosMaster, COL_CONTEXT)
        ' ProductID
            sProdId = GetProductId(sProductCode, iPosMaster) 
            If NOT sProdId = "" Then arrMaster(iPosMaster, COL_PRODUCTID) = sProdId
        ' ProductVersion
            arrMaster(iPosMaster, COL_PRODUCTVERSION) = GetProductVersion(sProductCode, iContext, sSid)
            If NOT fBasicMode Then
            ' cached .msi package 
                Set MsiDb = Nothing
                arrMaster(iPosMaster, COL_CACHEDMSI) = GetCachedMsi(sProductCode, iPosMaster) 
                If (NOT arrMaster(iPosMaster, COL_CACHEDMSI) = "") AND (arrMaster(iPosMaster, COL_VIRTUALIZED) = 0) Then
                    Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster, COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
                    arrMaster(iPosMaster, COL_PACKAGECODE) = GetPackageCode(sProductCode, iPosMaster, MsiDb)
                    arrMaster(iPosMaster, COL_UPGRADECODE) = GetUpgradeCode(MsiDb)
                    arrMaster(iPosMaster, COL_TRANSFORMS) = GetTransforms(sProductCode, iPosMaster)
                    ' InstallDate
                    arrMaster(iPosMaster, COL_INSTALLDATE) = GetInstallDate(sProductCode, iContext, sSid, arrMaster(iPosMaster, COL_CACHEDMSI))
                    arrMaster(iPosMaster, COL_ORIGINALMSI) = GetOriginalMsiName(sProductCode, iPosMaster)
                End If 'msi not empty
            End If 'fBasicMode
            
            If arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) Then
                ' SP level
                iVersionMajor = GetVersionMajor(sProductCode) 
                sSpLevel = OVersionToSpLevel(sProductCode, iVersionMajor, arrMaster(iPosMaster, COL_PRODUCTVERSION)) 
                arrMaster(iPosMaster, COL_SPLEVEL) = sSpLevel 
                ' Build/Origin
                If NOT fBasicMode Then arrMaster(iPosMaster, COL_ORIGIN) = CheckOrigin(MsiDb)
                ' Architecture (Bitness)
                If Left(arrMaster(iPosMaster, COL_PRODUCTVERSION), 2)>11 Then
                    If Mid(sProductCode, 21, 1) = "1" Then arrMaster(iPosMaster, COL_ARCHITECTURE) = "x64" Else arrMaster(iPosMaster, COL_ARCHITECTURE) = "x86"
                End If
                ' Key ComponentStates
                arrMaster(iPosMaster, COL_KEYCOMPONENTS) = GetKeyComponentStates(sProductCode, False)
                
                If NOT fBasicMode Then
                ' Cx
                    fVer = False : fCx = False
                    Select Case iVersionMajor
                    Case 11
                        sComp = "{1EBDE4BC-9A51-4630-B541-2561FA45CCC5}"
                        sRef  = "11.0.8320.0"
                    Case 12
                        sComp = "{0638C49D-BB8B-4CD1-B191-051E8F325736}"
                        sRef  = "12.0.6514.5001"
                        If Mid(sProductCode, 11, 4) = "0020" Then fVer = True
                    Case 14
                        If Mid(sProductCode, 21, 1) = "1" Then
                            sComp = "{C0AC079D-A84B-4CBD-8DBA-F1BB44146899}"
                            sComp2= "{E6AC97ED-6651-4C00-A8FE-790DB0485859}"
                        Else
                            sComp = "{019C826E-445A-4649-A5B0-0BF08FCC4EEE}"
                            sComp2= "{398E906A-826B-48DD-9791-549C649CACE5}"
                        End If
                        sRef  = "14.0.5123.5004"
                    Case Else
                        sComp = "" : sComp2 = "" : sRef = ""
                    End Select
                    ' obtain product handle
                    Err.Clear
                    If oMsi.Product(sProductCode, sSid, iContext).ComponentState(sComp) = INSTALLSTATE_LOCAL Then
                        If Err = 0 Then
                            fCx = True
                            sPath = oMsi.ComponentPath(sProductCode, sComp)
                            If oFso.FileExists(sPath) Then
                                fVer = (CompareVersion(oFso.GetFileVersion(sPath), sRef, True) > -1)
                            End If
                            If iVersionMajor = 14 Then
                                If oMsi.Product(sProductCode, sSid, iContext).ComponentState(sComp2) = INSTALLSTATE_LOCAL Then
                                    sPath = oMsi.ComponentPath(sProductCode, sComp2)
                                    If oFso.FileExists(sPath) Then
                                        fVer = (CompareVersion(oFso.GetFileVersion(sPath), sRef, True) > -1)
                                    End If
                                End If 'INSTALLSTATE_LOCAL
                                Err.Clear
                            End If
                        Else
                            Err.Clear
                        End If 'Err = 0 
                    End If 'INSTALLSTATE_LOCAL
                    If fVer Then
                        arrCxN = Array("44","43","58","46")
                        Select Case iVersionMajor
                        Case 11
                            arrCx = Array("53","4F","46","54","57","41","52","45","5C","4D","69","63","72","6F","73","6F","66","74","5C","4F","66","66","69","63","65","5C","31","31","2E","30","5C","57","6F","72","64")
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","4B","42","20","39","37","39","30","34","35")
                            If RegEnumValues(HKLM, hAtS(arrCx), arrNames, arrTypes) Then
                                For Each n in arrNames
                                    If UCase(n) = hAtS(arrCxN) Then arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & hAtS(arrCxErr) & CSV
                                Next 'n
                            End If
                        Case 12
                            arrCx = Array("53","4F","46","54","57","41","52","45","5C","4D","69","63","72","6F","73","6F","66","74","5C","4F","66","66","69","63","65","5C","31","32","2E","30","5C","57","6F","72","64")
                            arrCxErr = Array("57","6F","72","64","20","43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","4B","42","20","39","37","34","36","33","31")
                            If Mid(sProductCode, 11, 4) = "0020" Then
                                If CompareVersion(sRef, GetMsiProductVersion(arrMaster(iPosMaster, COL_CACHEDMSI)), False) < 1 Then
                                    arrCxErr = Array("57","6F","72","64","20","43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","64","69","73","61","62","6C","65","64","20","62","79","20","43","6F","6D","70","61","74","69","62","69","6C","69","74","79","20","50","61","63","6B")
                                End If
                            End If '"0020"
                            If RegEnumValues(HKLM, hAtS(arrCx), arrNames, arrTypes) Then
                                For Each n in arrNames
                                    If UCase(n) = hAtS(arrCxN) Then arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & hAtS(arrCxErr) & CSV
                                Next 'n
                            End If
                        Case 14
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","65","6E","61","62","6C","65","64","20","62","79","20","57","6F","72","64","20","32","30","31","30","20","4B","42","20","32","34","32","38","36","37","20","61","64","64","2D","69","6E")
                            For Each prod in arrMaster
                                If Mid(prod, 11, 4) = "0126" Then arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & hAtS(arrCxErr) & CSV
                            Next 'prod
                        Case Else
                        End Select
                    Else
                        If (iVersionMajor = 14) AND fCx Then
                            arrCxErr = Array("43","75","73","74","6F","6D","20","58","4D","4C","20","66","65","61","74","75","72","65","20","66","6F","72","20","74","68","65","20","62","69","6E","61","72","79","20","2E","64","6F","63","20","66","6F","72","6D","61","74","20","72","65","71","75","69","72","65","73","20","4B","42","20","32","34","31","33","36","35","39")
                            arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & hAtS(arrCxErr) & CSV
                        End If '14
                    End If 'fVer
                End If 'fBasicMode
            Else
                arrMaster(iPosMaster, COL_ORIGIN) = "n/a"
            ' checks for known add-ins
            ' POWERPIVOT_2010
                If InStr(POWERPIVOT_2010, arrMaster(iPosMaster, COL_UPGRADECODE)) > 0 Then
                    If CompareVersion(arrMaster(iPosMaster, COL_PRODUCTVERSION), "10.50.1747.0", True) < 1 Then
                        arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & ERR_CATEGORYWARN & "This is a preview version. Please obtain version 10.50.1747.0 or higher." & CSV
                    End If
                End If 'POWERPIVOT_2010
            End If 'IsOfficeProduct
        End If 'INSTALLSTATE_DEFAULT
    Next 'iPosMaster
End Sub 'ProductProperties
'=======================================================================================================

'The name of the original installation package 'PackageName' is obtained from HKCR.
Function GetOriginalMsiName(sProductCode, iPosMaster)
    Dim iPos
    Dim sCompGuid, sRegName
    Dim fVirtual
    On Error Resume Next

    fVirtual = (arrMaster(iPosMaster, COL_VIRTUALIZED) = 1)
    sRegName = ""
    If NOT fVirtual Then sRegName = oMsi.ProductInfo(sProductCode, "PackageName")

    'Error Handler
    If (Not Err = 0) OR sRegName = "" Then
        'This can happen if WI < 3.x or product is installed for other user
        Err.Clear
        iPos = GetArrayPosition(arrMaster, sProductCode)
        sCompGuid = GetCompressedGuid(sProductCode)
        sRegName = GetRegOriginalMsiName(sCompGuid, arrMaster(iPos, COL_CONTEXT), arrMaster(iPos, COL_USERSID))
        If sRegName = "-" AND NOT fVirtual Then arrMaster(iPos, COL_ERROR) = arrMaster(iPos, COL_ERROR) & ERR_CATEGORYERROR & ERR_BADMSINAMEMETADATA & CSV
    End If
    GetOriginalMsiName = sRegName
    
End Function
'=======================================================================================================

'The 'Transforms' property is obtained from HKCR.
Function GetTransforms(sProductCode, iPosMaster)
    Dim sTransforms, sCompGuid, sRegTransforms
    Dim iPos
    On Error Resume Next

    GetTransforms = "-" : sTransforms = ""
    If arrMaster(iPosMaster, COL_VIRTUALIZED) = 0 Then sTransforms = oMsi.ProductInfo(sProductCode, "Transforms")

    'Error Handler
    If NOT Err = 0 OR arrMaster(iPosMaster, COL_VIRTUALIZED) = 1 Then
        Err.Clear
        iPos = GetArrayPosition(arrMaster, sProductCode)
        sCompGuid = GetCompressedGuid(sProductCode)
        sTransforms = GetRegTransforms(sCompGuid, arrMaster(iPos, COL_CONTEXT), arrMaster(iPos, COL_USERSID))
    End If
    If Len(sTransforms) > 0 Then GetTransforms = sTransforms
End Function
'=======================================================================================================

'InstallDate is available as part of the ProductInfo.
Function GetInstallDate(sProductCode, iContext, sSid, sCachedMsi)
    Dim iPos
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sDateLocalized, sDateNormalized, sYY, sMM, sDD
    On Error Resume Next

    GetInstallDate = "" 
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, True) & "InstallProperties"
    sName = "InstallDate"
    GetInstallDate = "-"
    If RegReadValue(hDefKey, sSubKeyName, sName, sValue, "REG_EXPAND_SZ") Then GetInstallDate = sValue

    'The InstallDate is reset with every patch transaction
    'As a workaround the CreateDate of the cached .msi package will be used to obtain the correct date
    If oFso.FileExists(sCachedMsi) Then
        'GetInstallDate = oFso.GetFile(sCachedMsi).DateCreated
        sDateLocalized = oFso.GetFile(sCachedMsi).DateCreated
        sYY = Year(sDateLocalized)
        sMM = Right("0" & Month(sDateLocalized), 2)
        sDD = Right("0" & Day(sDateLocalized), 2)
        sDateNormalized = sYY & " " & sMM & " " & sDD & " (yyyy mm dd)"
        GetInstallDate = sDateNormalized
    End If
    
End Function 'GetInstallDate
'=======================================================================================================

'The package code associates a .msi file with an application or product
Function GetPackageCode(sProductCode, iPosMaster, MsiDb)
    Dim sValidate, sCompGuid, sPackageCode
    Dim oSumInfo
    On Error Resume Next

    sPackageCode = ""
    If arrMaster(iPosMaster, COL_VIRTUALIZED) = 0 Then sPackageCode = oMsi.ProductInfo(sProductCode, "PackageCode")
    If (Not Err = 0) OR (sPackageCode="") Then
    ' Error Handler
        sCompGuid = GetCompressedGuid(sProductCode)
        sPackageCode = GetRegPackageCode(sCompGuid, arrMaster(iPosMaster, COL_CONTEXT), arrMaster(iPosMaster, COL_USERSID))
        Exit Function
    End If
    If Not sPackageCode = "n/a" Then
        If Not IsValidGuid(sPackageCode, GUID_UNCOMPRESSED) Then
            If fGuidCaseWarningOnly Then
                arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & ERR_CATEGORYNOTE & ERR_GUIDCASE & DOT & sErrBpa & CSV
            Else
                Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & sProductCode & DSV & arrMaster(iPosMaster, COL_PRODUCTNAME) & _
                         ": " & sError & " for PackageCode '" & sPackageCode & "'" & DOT & sErrBpa
                sError = "" : sErrBpa = ""
                arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & sError & DOT & sErrBpa & CSV
            End If 'fGuidCaseWarningOnly
        End If
    End If
' Scan cached .msi
    If Not arrMaster(iPosMaster, COL_CACHEDMSI) = "" Then
	    Set oSumInfo = MsiDb.SummaryInformation(MSIOPENDATABASEMODE_READONLY)
	    If Not (Err = 0) Then
	        arrMaster(iPosMaster, COL_NOTES) = arrMaster(iPosMaster, COL_NOTES) & ERR_CATEGORYWARN & ERR_INITSUMINFO & CSV
            Exit Function
	    End If 'Not Err
	    If Not sPackageCode = oSumInfo.Property(PID_REVNUMBER) Then 
	        arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & ERR_PACKAGECODEMISMATCH & CSV
	        Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & sProductCode & DSV & arrMaster(iPosMaster, COL_PRODUCTNAME) & _
                     ": " & ERR_PACKAGECODEMISMATCH  & DOT & BPA_PACKAGECODEMISMATCH
        End If         
    End If 'arrMaster
    GetPackageCode = sPackageCode
End Function 'GetPackageCode
'=======================================================================================================

Function GetUpgradeCode (MsiDb)
    Dim Record
    Dim qView
    On Error Resume Next
    GetUpgradeCode = ""
    If MsiDb Is Nothing Then Exit Function
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property`='UpgradeCode'")
    qView.Execute()
    Set Record = qView.Fetch()
    If NOT Err = 0 Then Exit Function
    GetUpgradeCode = Record.StringData(1)
End Function 'GetPackageCode
'=======================================================================================================

'Read the Build / Origin property from the cached .msi
Function CheckOrigin(MsiDb)
    Dim sQuery, sCachedMsi
    Dim Record
    Dim qView
    On Error Resume Next

    CheckOrigin = ""
    If MsiDb Is Nothing Then Exit Function
' read the 'Build' entry first
	sQuery = "SELECT `Value` FROM Property WHERE `Property` = 'BUILD'"
	Set qView = MsiDb.OpenView(sQuery)
	qView.Execute 
	Set Record = qView.Fetch()
	If Not Record Is Nothing Then
	    CheckOrigin = Record.StringData(1)
	End If 'Is Nothing
	sQuery = "SELECT `Value` FROM Property WHERE `Property` = 'ORIGIN'"
	Set qView = MsiDb.OpenView(sQuery)
	qView.Execute 
	Set Record = qView.Fetch()
	If Not Record Is Nothing Then
	    CheckOrigin = CheckOrigin & " / " & Record.StringData(1)
	End If 'Is Nothing
End Function
'=======================================================================================================

Function GetCachedMsi(sProductCode, iPosMaster)
    Dim sCachedMsi
    Dim oApp
    Dim fVirtual
    On Error Resume Next

    fVirtual = (arrMaster(iPosMaster, COL_VIRTUALIZED) = 1)
    sCachedMsi = ""
    If NOT fVirtual Then
        If iWiVersionMajor > 2 Then
            Set oApp = oMsi.Product(sProductCode, arrMaster(iPosMaster, COL_USERSID), arrMaster(iPosMaster, COL_CONTEXT))
            sCachedMsi = oApp.InstallProperty("LocalPackage")
        ElseIf (arrMaster(iPosMaster, COL_USERSID) = sCurUserSid) Or (arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
            sCachedMsi = oMsi.ProductInfo(sProductCode, "LocalPackage")
        Else
        ' rely on error handling to retain the value from direct registry read
        End If 'iWiVersionMajor
    End If 'fVirtual

    If Not Err = 0 Then
        sCachedMsi= ""
        Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & sProductCode & DSV & arrMaster(iPosMaster, COL_PRODUCTNAME) & _
                 ": " & ERR_PACKAGEAPIFAILURE & "."
    End If
    
    If sCachedMsi= "" Then
        sCachedMsi = GetRegCachedMsi(GetCompressedGuid(sProductCode), iPosMaster)
    End If 'sCachedMsi= ""
    If sCachedMsi= "" Then
        If NOT fVirtual Then arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & ERR_BADPACKAGEMETADATA & CSV
    Else
        If Not oFso.FileExists (sCachedMsi) Then 
            Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & sProductCode & " - " &_
                     arrMaster(iPosMaster, COL_PRODUCTNAME) & ": " & ERR_LOCALPACKAGEMISSING & ": " & sCachedMsi
            arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & ERR_LOCALPACKAGEMISSING & CSV
            sCachedMsi = ""
        End If 'oFso.FileExists
    End If
    GetCachedMsi = sCachedMsi
End Function 'GetCachedMsi
'=======================================================================================================

Function GetRegCachedMsi (sProductCodeCompressed, iPosMaster)
    Dim hDefKey
    Dim sSubKeyName, sValue, sSid, sName
    Dim iContext
    On Error Resume Next

    GetRegCachedMsi = ""
    'Go global
    hDefKey = HKLM
    iContext = arrMaster(iPosMaster, COL_CONTEXT)
    sSid = arrMaster(iPosMaster, COL_USERSID)
    sName = "LocalPackage"
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ManagedLocalPackage"
    End If
    
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, True) & "InstallProperties\"
    If RegReadStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegCachedMsi = sValue
End Function
'=======================================================================================================

Function GetProductId(sProductCode, iPosMaster)
    Dim sProductId
    Dim oApp
    On Error Resume Next

    GetProductId = ""
    
    If arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_C2RV2 Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode), iPosMaster)
        Exit Function
    End If
    If arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_C2RV3 Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode), iPosMaster)
        Exit Function
    End If

    If iWiVersionMajor > 2 Then
        Set oApp = oMsi.Product(sProductCode, arrMaster(iPosMaster, COL_USERSID), arrMaster(iPosMaster, COL_CONTEXT))
        sProductId = oApp.InstallProperty("ProductID")
    ElseIf (arrMaster(iPosMaster, COL_USERSID) = sCurUserSid) Or (arrMaster(iPosMaster, COL_CONTEXT) = MSIINSTALLCONTEXT_MACHINE) Then
        sProductId = oMsi.ProductInfo(sProductCode, "ProductID")
    Else
        'Rely on error handling to retain the value from direct registry read
    End If 'iWiVersionMajor

    If Not Err = 0 Then
        sProductId= ""
    End If
    
    If sProductId= "" Then
        sProductId = GetRegProductId(GetCompressedGuid(sProductCode), iPosMaster)
    End If 'sCachedMsi= ""

    GetProductId = sProductId
End Function 'GetProductId
'=======================================================================================================

Function GetRegProductId (sProductCodeCompressed, iPosMaster)
    Dim hDefKey
    Dim sSubKeyName, sValue, sSid, sName
    Dim iContext
    On Error Resume Next

    GetRegProductId = ""
    'Go global
    hDefKey = HKLM
    iContext = arrMaster(iPosMaster, COL_CONTEXT)
    sSid = arrMaster(iPosMaster, COL_USERSID)
    sName = "ProductID"
    'Tweak managed to unmanaged to avoid link to managed global key
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ProductID"
    End If
    
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, True) & "InstallProperties\"
    If RegReadStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegProductId = sValue

End Function 'GetRegProductId
'=======================================================================================================

'Get the ProductVersion string from WI ProductInfo
Function GetProductVersion (sProductCode, iContext, sSid)
    Dim sTmp
    On Error Resume Next

    If iContext = MSIINSTALLCONTEXT_C2RV2 Then
         GetProductVersion = GetRegProductVersion(sProductCode, iContext, sSid)
         Exit Function
    End If
    If iContext = MSIINSTALLCONTEXT_C2RV3 Then
         GetProductVersion = GetRegProductVersion(sProductCode, iContext, sSid)
         Exit Function
    End If
    
    sTmp = ""
    sTmp = oMsi.ProductInfo (sProductCode, "VersionString")
    If (sTmp = "") OR (NOT Err = 0) Then
        Err.Clear
        sTmp = GetRegProductVersion(sProductCode, iContext, sSid)
    End If
    GetProductVersion = sTmp

End Function
'=======================================================================================================

'Get the ProductVersion from Registry
Function GetRegProductVersion (sProductCode, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sValue
    Dim iTmpContext
    On Error Resume Next

    GetRegProductVersion = "Error"
    hDefKey = HKEY_LOCAL_MACHINE
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iTmpContext = MSIINSTALLCONTEXT_USERUNMANAGED
    Else
        iTmpContext = iContext
    End If 'iContext = MSIINSTALLCONTEXT_USERMANAGED
    sSubKeyName = GetRegConfigKey(GetCompressedGuid(sProductCode), iTmpContext, sSid, True) & "InstallProperties\"
    If RegReadStringValue(hDefKey, sSubKeyName, "DisplayVersion", sValue) Then GetRegProductVersion = sValue
   
End Function
'=======================================================================================================

'Translate the Office ProductVersion to the service pack level
Function OVersionToSpLevel (sProductCode, iVersionMajor, sProductVersion)
    On Error Resume Next
    
    'SKU identifier constants for SP level detection
    Const O16_EXCEPTION = ""
    Const O15_EXCEPTION = ""
    Const O14_EXCEPTION = "007A, 007B, 007C, 007D, 007F, 2005"
'   O12_Server = "1014, 1015, 104B, 104E, 1080, 1088, 10D7, 10D8, 10EB, 10F5, 10F6, 10F7, 10F8, 10FB, 10FC, 10FD, 1103, 1104, 110D, 1105, 1110, 1121, 1122"    '#Devonly
    Const O12_EXCEPTION = "001C, 001F, 0020, 003F, 0045, 00A4, 00A7, 00B0, 00B1, 00B2, 00B9, 011F, CFDA"
    Const O11_EXCEPTION = "14, 15, 16, 17, 18, 19, 1A, 1B, 1C, 24, 32, 3A, 3B, 44, 51, 52, 53, 5E, A1, A4, A9, E0"
    Const O10_EXCEPTION = "17, 1D, 25, 27, 30, 36, 3A, 3B, 51, 52, 53, 54"
    Const O09_EXCEPTION = "3A, 3B, 3C, 5F"

    Dim iSpCnt, iExptnCnt, iLevel, iRetry
    Dim sSpLevel, sSku

    iLevel = 0 : iRetry = 0
    Select Case iVersionMajor

    Case 9
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid (sProductCode, 4, 2)
        If InStr (O09_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound (arrProdVer09, 1)
                If InStr (arrProdVer10 (iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O09_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound (arrProdVer09, 2)
                If sProductVersion = Left (arrProdVer09 (iExptnCnt, iSpCnt), Len (sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr (arrProdVer09 (iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid (arrProdVer09 (iExptnCnt, iSpCnt), InStr (arrProdVer09 (iExptnCnt, iSpCnt), ",") + 1, Len (arrProdVer09 (iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 10
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid(sProductCode, 4, 2)
        If InStr(O10_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer10, 1)
                If InStr(arrProdVer10(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O10_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer10, 2)
                If sProductVersion = Left(arrProdVer10(iExptnCnt, iSpCnt), Len(sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer10(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer10(iExptnCnt, iSpCnt), InStr(arrProdVer10(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer10(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 11
        'Sku ProductID is 2 digits starting at pos 4
        sSku = Mid(sProductCode, 4, 2)
        If InStr(O11_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer11, 1)
                If InStr(arrProdVer11(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O11_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer11, 2)
                If sProductVersion = Left(arrProdVer11(iExptnCnt, iSpCnt), Len(sProductVersion)) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer11(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer11(iExptnCnt, iSpCnt), InStr(arrProdVer11(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer11(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 12
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O12_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 2 To UBound(arrProdVer12, 1)
                If InStr(arrProdVer12(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        ElseIf Left(sSku, 1) = "1" Then 'Server SKU
            iExptnCnt = 1
        Else
            iExptnCnt = 0
        End If 'InStr(O12_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer12, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer12(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer12(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer12(iExptnCnt, iSpCnt), InStr(arrProdVer12(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer12(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 14
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O14_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer14, 1)
                If InStr(arrProdVer14(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O14_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer14, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer14(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer14(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer14(iExptnCnt, iSpCnt), InStr(arrProdVer14(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer14(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry

    Case 15
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O15_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer15, 1)
                If InStr(arrProdVer15(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O15_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer15, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer15(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer15(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer15(iExptnCnt, iSpCnt), InStr(arrProdVer15(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer15(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry
        
        Case 16
        'Sku ProductID is 4 digits starting at pos 11
        sSku = Mid(sProductCode, 11, 4)
        If InStr(O16_EXCEPTION, sSku) > 0 Then
            For iExptnCnt = 1 To UBound(arrProdVer16, 1)
                If InStr(arrProdVer16(iExptnCnt, 0), sSku) > 0 Then Exit For
            Next 'iExptnCnt
        Else
            iExptnCnt = 0
        End If 'InStr(O16_Exception, sSku) > 0
        
        For iRetry = 0 To 1
            For iSpCnt = 1 To UBound(arrProdVer16, 2)
                If Left(sProductVersion, 10) = Left(arrProdVer16(iExptnCnt, iSpCnt), 10) Then 
                    'Special release references are noted within same field with a "," separator
                    If InStr(arrProdVer16(iExptnCnt, iSpCnt), ",") > 0 Then
                        OVersionToSpLevel = Mid(arrProdVer16(iExptnCnt, iSpCnt), InStr(arrProdVer16(iExptnCnt, iSpCnt), ",") + 1, Len(arrProdVer16(iExptnCnt, iSpCnt)))
                        Exit Function
                    Else
                        iLevel = iSpCnt
                        Exit For
                    End If
                End If
            Next 'iSpCnt
            If iLevel > 0 Then Exit For
            'Did not find the SP level yet. Retry with core build numbers
            iExptnCnt = 0
        Next 'iRetry
Case Else
    End Select
        
    Select Case iLevel
    Case 1 : sSpLevel = "RTM"
    Case 2 : sSpLevel = "SP1"
    Case 3 : sSpLevel = "SP2"
    Case 4 : sSpLevel = "SP3"
    Case Else : sSpLevel = ""
    End Select

    OVersionToSpLevel = sSpLevel
    
End Function 'OVersionToSpLevel
'=======================================================================================================

'Initialize arrays for translation ProductVersion -> ServicePackLevel
Sub InitProdVerArrays
    On Error Resume Next

' O16 Products -> KB ?
    ReDim arrProdVer16(0, 0) 'n, 1=RTM
    arrProdVer16(0, 0) = "" 

' 2013 Products -> KB 2786054
    ReDim arrProdVer15(0, 44) 'n, 1=RTM
    arrProdVer15(0, 0) = "" : arrProdVer15(0, 1) = "15.0.4420.1017" : arrProdVer15(0, 2) = "15.0.4569.1507"
    arrProdVer15(0, 3) = "15.0.4454.1004, 2013/01" : arrProdVer15(0, 4) = "15.0.4454.1511, 2013/02" 
    arrProdVer15(0, 5) = "15.0.4481.1005, 2013/03" : arrProdVer15(0, 6) = "15.0.4481.1510, 2013/04" : arrProdVer15(0, 7) = "15.0.4505.1006, 2013/05" 
    arrProdVer15(0, 8) = "15.0.4505.1510, 2013/06" : arrProdVer15(0, 8) = "15.0.4517.1005, 2013/07" : arrProdVer15(0, 9) = "15.0.4517.1509, 2013/08"
    arrProdVer15(0, 10) = "15.0.4535.1004, 2013/09" : arrProdVer15(0, 11) = "15.0.4535.1511, 2013/10" : arrProdVer15(0, 12) = "15.0.4551.1005, 2013/11"
    arrProdVer15(0, 13) = "15.0.4551.1011, 2013/12" : arrProdVer15(0, 14) = "15.0.4551.1512, 2014/01"
    arrProdVer15(0, 15) = "15.0.4569.1000, SP1_Preview" : arrProdVer15(0, 16) = "15.0.4569.1001, SP1_Preview" : arrProdVer15(0, 17) = "15.0.4569.1002, SP1_Preview" : arrProdVer15(0, 18) = "15.0.4569.1003, SP1_Preview"
    arrProdVer15(0, 19) = "15.0.4569.1508, 2014/03" : arrProdVer15(0, 20) = "15.0.4605.1003, 2014/04" : arrProdVer15(0, 21) = "15.0.4615.1002, 2014/05" : arrProdVer15(0, 22) = "15.0.4623.1003, 2014/06"
    arrProdVer15(0, 23) = "15.0.4631.1002, 2014/07" : arrProdVer15(0, 24) = "15.0.4631.1004, 2014/07+" : arrProdVer15(0, 25) = "15.0.4641.1002, 2014/08" : arrProdVer15(0, 26) = "15.0.4641.1003, 2014/08+"
    arrProdVer15(0, 27) = "15.0.4649.1001, 2014/09" : arrProdVer15(0, 28) = "15.0.4649.1001, 2014/09+" : arrProdVer15(0, 29) = "15.0.4649.1004, 2014/09++" : arrProdVer15(0, 30) = "15.0.4659.1001, 2014/10"
    arrProdVer15(0, 31) = "15.0.4667.1002, 2014/11" : arrProdVer15(0, 32) = "15.0.4675.1002, 2014/12" : arrProdVer15(0, 33) = "15.0.4675.1003, 2014/12+" : arrProdVer15(0, 34) = "15.0.4693.1001, 2015/02"
    arrProdVer15(0, 35) = "15.0.4693.1002, 2015/02+" : arrProdVer15(0, 36) = "15.0.4701.1002, 2015/03" : arrProdVer15(0, 37) = "15.0.4711.1002, 2015/04" : arrProdVer15(0, 38) = "15.0.4711.1003, 2015/04+"
    arrProdVer15(0, 39) = "15.0.4719.1002, 2015/05" : arrProdVer15(0, 40) = "15.0.4727.1002, 2015/06" : arrProdVer15(0, 41) = "15.0.4727.1003, 2015/06+" : arrProdVer15(0, 42) = "15.0.4737.1003, 2015/07"
    arrProdVer15(0, 43) = "15.0.4745.1001, 2015/08" : arrProdVer15(0, 44) = "15.0.4745.1001, 2015/08+"

' 2010 Products -> KB 2186281
    ReDim arrProdVer14(6, 3) 'n, 1=RTM; n, 2=SP1
    arrProdVer14(0, 0) = "" : arrProdVer14(0, 1) = "14.0.4763.1000" : arrProdVer14(0, 2) = "14.0.6029.1000" : arrProdVer14(0, 3) = "14.0.7015.1000"
    arrProdVer14(1, 0) = "007A" : arrProdVer14(1, 1) = "14.0.5118.5000, Web V1" 'Outlook Connector
    arrProdVer14(2, 0) = "007B" : arrProdVer14(2, 1) = "14.0.5117.5000, Web V1" 'Outlook Social Connector Provider for Windows Live Messen
    arrProdVer14(3, 0) = "007C" : arrProdVer14(3, 1) = "14.0.5117.5000, Web V1" 'Outlook Social Connector Facebook
    arrProdVer14(4, 0) = "007D" : arrProdVer14(4, 1) = "14.0.5120.5000, Web V1" 'Outlook Social Connector Windows Live
    arrProdVer14(5, 0) = "007F" : arrProdVer14(5, 1) = "14.0.5139.5001, Web V1" 'Outlook Connector
    arrProdVer14(6, 0) = "2005" : arrProdVer14(6, 1) = "14.0.5130.5003, Web V1" 'OFV File Validation Add-In
    
' 2007 Products -> KB 928516
    ReDim arrProdVer12(12, 4) 'n, 1=RTM; n, 2=SP1
    arrProdVer12(0, 0) = "": arrProdVer12(0, 1) = "12.0.4518.1014": arrProdVer12(0, 2) = "12.0.6215.1000" : arrProdVer12(0, 3) = "12.0.6425.1000" : arrProdVer12(0, 4) = "12.0.6612.1000"
    arrProdVer12(1, 0) = "" : arrProdVer12(1, 1) = "12.0.4518.1016": arrProdVer12(1, 2) = "12.0.6219.1000" ': arrProdVer12(1, 3) = "12.0.6425.1000" 'Server
    arrProdVer12(2, 0) = "0045" : arrProdVer12(2, 1) = "12.0.4518.1084" 'Expression Web 2
    arrProdVer12(3, 0) = "011F" : arrProdVer12(3, 1) = "12.0.6407.1000, Web V1 (Windows Live)" 'Outlook Connector
    arrProdVer12(4, 0) = "001F" : arrProdVer12(4, 1) = "12.0.4518.1014" : arrProdVer12(4, 2) = "12.0.6213.1000" 'Office Proof (Container)
    arrProdVer12(5, 0) = "00B9" : arrProdVer12(5, 1) = "12.0.6012.5000, KB 932080": arrProdVer12(5, 2) = "12.0.6015.5000, Web V1 (Windows Live)"'Application Error Reporting
    arrProdVer12(6, 0) = "0020" : arrProdVer12(6, 1) = "12.0.4518.1014, Web V1" : arrProdVer12(6, 3) = "12.0.6021.5000, Web V3" : arrProdVer12(6, 4) = "12.0.6514.5001, Web V4 (incl. SP2)"  ': arrProdVer12(6, 5) = "12.0.6425.1000, SP2"'Compatibility Pack
    arrProdVer12(7, 0) = "00B0, 00B1, 00B2" : arrProdVer12(7, 1) = "12.0.4518.1014, Web V1" 'Save As packs
    arrProdVer12(8, 0) = "004A" : arrProdVer12(8, 1) = "12.0.4518.1014, Web RTM" : arrProdVer12(8, 2) = "12.0.6213.1000, Web SP1" 'Web Components
    arrProdVer12(9, 0) = "00A7" : arrProdVer12(9, 1) = "12.0.4518.1014, Web RTM" : arrProdVer12(9, 2) = "12.0.6520.3001, SP2" 'Calendar Printing Assistant
    arrProdVer12(10, 0) = "CFDA" : arrProdVer12(10, 1) = "12.0.613.1000, RTM" : arrProdVer12(10, 2) = "12.1.1313.1000, SP1"  : arrProdVer12(10, 2) = "12.0.6511.5000, SP2" 'Project Portfolio Server 2007
    arrProdVer12(11, 0) = "001C" : arrProdVer12(11, 1) = "12.0.4518.1049, RTM"  : arrProdVer12(11, 2) = "12.0.6230.1000, SP1" : arrProdVer12(11, 3) = "12.0.6237.1003, SP1" 'Access Runtime
    arrProdVer12(12, 0) = "003F" : arrProdVer12(12, 1) = "12.0.6214.1000, SP1" 'Excel Viewer

' 2003 Products -> KB 832672
    'ProductVersions      -> KB821549
    Redim arrProdVer11(16, 4) 'n, 1=RTM; n, 2=SP1; n, 3=SP2; n, 4=SP3
    arrProdVer11(0, 0) = "": arrProdVer11(0, 1) = "11.0.5614.0": arrProdVer11(0, 2) = "11.0.6361.0": arrProdVer11(0, 3) = "11.0.7969.0": arrProdVer11(0, 4) = "11.0.8173.0" 'Suites & Core
    arrProdVer11(1, 0) = "14": arrProdVer11(1, 1) = "11.0.5614.0": arrProdVer11(1, 2) = "11.0.6361.0": arrProdVer11(1, 3) = "11.0.7969.0": arrProdVer11(1, 4) = "11.0.8173.0" 'Sharepoint Services
    arrProdVer11(2, 0) = "15, 1C": arrProdVer11(2, 1) = "11.0.5614.0": arrProdVer11(2, 2) = "11.0.6355.0"': arrProdVer11(2, 3) = "11.0.7969.0": arrProdVer11(2, 4) = "11.0.8173.0" 'Access
    arrProdVer11(3, 0) = "16": arrProdVer11(3, 1) = "11.0.5612.0": arrProdVer11(3, 2) = "11.0.6355.0"': arrProdVer11(3, 3) = "11.0.7969.0": arrProdVer11(3, 4) = "11.0.8173.0" 'Excel
    arrProdVer11(4, 0) = "17": arrProdVer11(4, 1) = "11.0.5516.0": arrProdVer11(4, 2) = "11.0.6356.0"': arrProdVer11(4, 3) = "11.0.7969.0": arrProdVer11(4, 4) = "11.0.8173.0" 'FrontPage
    arrProdVer11(5, 0) = "18": arrProdVer11(5, 1) = "11.0.5529.0": arrProdVer11(5, 2) = "11.0.6361.0"': arrProdVer11(5, 3) = "11.0.7969.0": arrProdVer11(5, 4) = "11.0.8173.0" 'PowerPoint
    arrProdVer11(6, 0) = "19": arrProdVer11(6, 1) = "11.0.5525.0": arrProdVer11(6, 2) = "11.0.6255.0"': arrProdVer11(6, 3) = "11.0.7969.0": arrProdVer11(6, 4) = "11.0.8173.0" 'Publisher
    arrProdVer11(7, 0) = "1A, E0": arrProdVer11(7, 1) = "11.0.5510.0": arrProdVer11(7, 2) = "11.0.6353.0"': arrProdVer11(7, 3) = "11.0.7969.0": arrProdVer11(7, 4) = "11.0.8173.0" 'Outlook
    arrProdVer11(8, 0) = "1B": arrProdVer11(8, 1) = "11.0.5510.0": arrProdVer11(8, 2) = "11.0.6353.0"': arrProdVer11(8, 3) = "11.0.7969.0": arrProdVer11(8, 4) = "11.0.8173.0" 'Word
    arrProdVer11(9, 0) = "44": arrProdVer11(9, 1) = "11.0.5531.0": arrProdVer11(9, 2) = "11.0.6357.0"': arrProdVer11(9, 3) = "11.0.7969.0": arrProdVer11(9, 4) = "11.0.8173.0" 'InfoPath
    arrProdVer11(10, 0) = "A1": arrProdVer11(10, 1) = "11.0.5614.0": arrProdVer11(10, 2) = "11.0.6360.0"': arrProdVer11(10, 3) = "11.0.7969.0": arrProdVer11(10, 4) = "11.0.8173.0" 'OneNote
    arrProdVer11(11, 0) = "3A, 3B, 32": arrProdVer11(11, 1) = "11.0.5614.0": arrProdVer11(11, 2) = "11.0.6707.0"': arrProdVer11(11, 3) = "11.0.7969.0": arrProdVer11(11, 4) = "11.0.8173.0" 'Project
    arrProdVer11(12, 0) = "51, 53, 5E": arrProdVer11(12, 1) = "11.0.3216.5614": arrProdVer11(12, 2) = "11.0.4301.6360"': arrProdVer11(12, 3) = "11.0.7969.0": arrProdVer11(12, 4) = "11.0.8173.0" 'Visio
    arrProdVer11(13, 0) = "24" : arrProdVer11(13, 1) = "11.0.5614.0, V1 (Web)" : arrProdVer11(13, 2) = "11.0.6120.0, V2 (Hotfix)": arrProdVer11(13, 3) = "11.0.6550.0, V2 (Japanese Hotfix)" 'ORK
    arrProdVer11(14, 0) = "52": arrProdVer11(14, 1) = "11.0.3709.5614": arrProdVer11(14, 2) = "11.0.6206.8011, Web V1"': arrProdVer11(14, 3) = "11.0.7969.0": arrProdVer11(14, 4) = "11.0.8173.0" 'Visio Viewer
    arrProdVer11(15, 0) = "A4" : arrProdVer11(15, 1) = "11.0.5614.0": arrProdVer11(15, 2) = "11.0.6361.0" : arrProdVer11(15, 3) = "11.0.6558.0, MSDE"': arrProdVer11(15, 4) = "11.0.7969.0, SP2": arrProdVer11(15, 5) = "11.0.8173.0, SP3"'OWC11
    arrProdVer11(16, 0) = "A9" : arrProdVer11(16, 1) = "11.0.5614.0, RTM" : arrProdVer11(16, 2) = "11.0.6553.0, Web V1" 'PIA11

' Office XP
    'Office 10 Numbering Scheme -> KB302663
    'ProuctVersions             -> KB291331
    Redim arrProdVer10(5, 4)
    arrProdVer10(0, 0) = "": arrProdVer10(0, 1) = "10.0.2627.01": arrProdVer10(0, 2) = "10.0.3520.0": arrProdVer10(0, 3) = "10.0.4330.0": arrProdVer10(0, 4) = "10.0.6626.0"
    'LPK/MUI RTM have inconsisten RTM build versions but do follow the main SP builds.
    '=> Limitation to not cover RTM LPK's
    arrProdVer10(1, 0) = "17, 1D": arrProdVer10(1, 1) = "10.0.2623.0": arrProdVer10(1, 2) = "10.0.3506.0": arrProdVer10(1, 3) = "10.0.4128.0": arrProdVer10(1, 4) = "10.0.6308.0"'FrontPage
    arrProdVer10(2, 0) = "27, 3A, 3B": arrProdVer10(2, 1) = "10.0.2915.0": arrProdVer10(2, 2) = "10.0.3416.0": arrProdVer10(2, 3) = "10.0.4219.0": arrProdVer10(2, 4) = "10.0.6612.0"'Project
    arrProdVer10(3, 0) = "30": arrProdVer10(3, 1) = "10.0.2619.0" 'Media Content (CAG)
    arrProdVer10(4, 0) = "25, 36": arrProdVer10(4, 1) = "10.0.2701.0, V1": arrProdVer10(4, 2) = "10.0.6403.0, V2 (Oct 24 2002 - Release)" 'ORK
    arrProdVer10(5, 0) = "51, 52, 53, 54": arrProdVer10(5, 1) = "10.0.525": arrProdVer10(5, 2) = "10.1.2514" : arrProdVer10(5, 3) = "10.2.5110" 'Visio

' 2000
    Redim arrProdVer09(2, 4)
    arrProdVer09(0, 0) = "": arrProdVer09(0, 1) = "9.00.2720": arrProdVer09(0, 2) = "9.00.3821, SR1": arrProdVer09(0, 3) = "9.00.4527": arrProdVer09(0, 4) = "9.00.9327"
    arrProdVer09(1, 0) = "3A, 3B, 3C": arrProdVer09(1, 1) = "9.00.2720": arrProdVer09(1, 2) = "9.00.4527"
    arrProdVer09(2, 0) = "5F":arrProdVer09(2, 1) = "9.00.00.2010, Web V1" 'ORK

    fInitArrProdVer = True
End Sub

'=======================================================================================================
'Module Patch 
'=======================================================================================================
Sub FindAllPatches
    If fBasicMode Then Exit Sub
    Dim n
    On Error Resume Next
    
    'Set Default for patch array
    Redim arrPatch(UBound(arrMaster), PATCH_COLUMNCOUNT, 0)
    
    'Check which patch API calls are safe for use
    CheckPatchApi
    
    'Iterate all products to add the patches to the array 
    For n = 0 To UBound(arrMaster)
        FindPatches(n)
    Next 'n
    
    AddDetailsFromPackage
End Sub
'=======================================================================================================

'Find all patches registered to a product and add them to the Patch array
Sub FindPatches (iPosMaster)
    On Error Resume Next
    
    Dim Patches, Patch, siSumInfo, MsiDb, Record, qView
    Dim sQuery, sErr
    Dim iMspMax, iMspCnt
    Dim fVirtual

    fVirtual = (arrMaster(iPosMaster, COL_VIRTUALIZED) = 1)
    'Find client side patches (CSP) for the product
    If iWiVersionMajor > 2 Then
        If NOT fVirtual Then Set Patches = oMsi.PatchesEx(arrMaster(iPosMaster, COL_PRODUCTCODE), arrMaster(iPosMaster, COL_USERSID), arrMaster(iPosMaster, COL_CONTEXT), MSIPATCHSTATE_ALL)
        If Not Err = 0 OR fVirtual Then
            'PatchesEx API call failed
            'Log the error
            sErr = GetErrorDescription(Err)
            If fPatchesExOk AND NOT fVirtual Then
                'Only log the error if fPatchesExOk = True.
                'If we're Not fPatchesExOk the error is expected and has already been logged.
                Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Product " & arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster, COL_PRODUCTNAME) & _
                     ": " & ERR_PATCHESEX
                arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & ERR_CATEGORYERROR & ERR_PATCHESEX & DSV & sErr & CSV
            End If 'fPatchesExOk
            'Fall back to registry detection
            FindRegPatches iPosMaster
        Else
            'PatchesEx API call succeeded
            'Ensure sufficient array size
            iMspMax = Patches.Count
            If iMspMax > UBound(arrPatch, 3) Then ReDim Preserve arrPatch(UBound(arrMaster), PATCH_COLUMNCOUNT, iMspMax)
            'Find the patch details
            ReadPatchDetails Patches, iPosMaster
        End If 'Err = 0
        
    Else
        FindRegPatches iPosMaster
    End If 'iWiVersionMajor > 2
    
    'Find patches in InstallSource (AIP)
    'Only valid for Office products
    Err.Clear
    If arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) Then
        Set MsiDb = oMsi.OpenDatabase(arrMaster(iPosMaster, COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
        'Critical error for this sub. Need to terminate it to prevent endless loop. '#Devonly
        If Not Err = 0 Then
            Err.Clear
            'Error has already been logged '#Devonly
            Exit Sub
        End If

        sQuery = "SELECT * FROM Property"
        Set qView = MsiDb.OpenView(sQuery)
        qView.Execute 
        Set Record = qView.Fetch()
        ' Loop through the records in the view
        Do Until (Record Is Nothing)
            'Get the properties name. The 'OR' condition covers the  Office 2000 family AIP's
            If (Left(Record.StringData(1), 1) = "_" And Mid(Record.StringData(1), 15, 1) = "_") _
            OR (Left(Record.StringData(1), 1) = "{" And Mid(Record.StringData(1), 15, 1) = "-") Then 
		        'Found AIP patch
		        'ReDim patch array '#Devonly
		        iMspCnt = UBound(arrPatch, 3) + 1
		        ReDim Preserve arrPatch(UBound(arrMaster), PATCH_COLUMNCOUNT, iMspCnt)
		        'Client side patch flag
		        arrPatch(iPosMaster, PATCH_CSP, iMspCnt) = False
	            'PatchCode
	            arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) = _
	            "{" & Mid(Record.StringData(1), 2, 8) & _
	            "-" & Mid(Record.StringData(1), 11, 4) & _
	            "-" & Mid(Record.StringData(1), 16, 4) & _
	            "-" & Mid(Record.StringData(1), 21, 4) & _
	            "-" & Mid(Record.StringData(1), 26, 12) &  "}"
	            'DisplayName
	            arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt) = Record.StringData(2)
            End If 
            ' Next record
            Set Record = qView.Fetch()
        Loop
    End If 'arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)

End Sub 'FindPatches
'=======================================================================================================

'Find all patches registered to a product by direct registry read and add them to the Patch array
Sub FindRegPatches (iPosMaster)
    Dim sSid, sSubKeyName, sSubKeyNamePatch, sSubKeyNamePatches, sValue, sProductCode
    Dim sPatch, sPatches, sAllPatches, sTmpPatches
    Dim hDefKey, hDefKeyGlobal
    Dim bPatchRegOk, bPatchMultiSzBroken, bPatchOrgMspNameBroken, bPatchCachedMspLinkBroken
    Dim bPatchGlobalOK, bPatchProdGlobalBroken, bPatchGlobalMultiSzBroken
    Dim iContext, iName, iType, iKey, iState, iMspMax
    Dim arrName, arrType, arrKeys, arrRegPatch
    Const REGDIM1 = 1
    On Error Resume Next
    
    iContext = arrMaster(iPosMaster, COL_CONTEXT)
    sSid = arrMaster(iPosMaster, COL_USERSID)
    sProductCode = arrMaster(iPosMaster, COL_PRODUCTCODE)
    sTmpPatches = ""
    bPatchRegOk = False
    bPatchGlobalOk = False
    bPatchMultiSzBroken = False
    bPatchOrgMspNameBroken = False
    ReDim arrRegPatch(REGDIM1, - 1)
    
    'Find "Applied" patches
    hDefKeyGlobal = GetRegHive(iContext, sSid, True)
    hDefKey = GetRegHive(iContext, sSid, False)
    sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, False) & "Patches\"
    If RegEnumValues(hDefKey, sSubKeyName, arrName, arrType) AND CheckArray(arrName) Then
        bPatchRegOk = True
        'Check Metadata integrity
        sPatches = vbNullString
        If Not RegReadValue(hDefKey, sSubKeyName, "Patches", sPatches, REG_MULTI_SZ) Then bPatchRegOk = False
        sTmpPatches = sPatches
        For iName = 0 To UBound(arrName)
            If arrType(iName) = REG_SZ Then
                ReDim Preserve arrRegPatch(REGDIM1, UBound(arrRegPatch, 2) + 1)
                sValue = ""
                If RegReadStringValue(hDefKey, sSubKeyName, arrName(iName), sValue) Then sPatch = arrName(iName)
                arrRegPatch(0, UBound(arrRegPatch, 2)) = GetExpandedGuid(sPatch)
                arrRegPatch(1, UBound(arrRegPatch, 2)) = MSIPATCHSTATE_APPLIED
                If NOT InStr(sTmpPatches, sPatch) > 0 Then
                    bPatchRegOk = False
                    bPatchMultiSzBroken = True
                Else
                    'Strip current patch from sTmpPatches
                    sTmpPatches = Replace(sTmpPatches, sPatch, "")
                End If 'InStr(sTmpPatches, sPatch)
                
                'Check Patches key
                sSubKeyNamePatches = GetRegConfigPatchesKey(iContext, sSid, False)
                If Not RegKeyExists(hDefKey, sSubKeyNamePatches & sPatch) Then
                    bPatchOrgMspNameBroken = True
                End If
                
                'Check Global Patches key
                sSubKeyNamePatches = ""
                sSubKeyNamePatches = GetRegConfigPatchesKey(iContext, sSid, True)
                If Not RegKeyExists(hDefKeyGlobal, sSubKeyNamePatches & sPatch) Then
                    bPatchCachedMspLinkBroken = True
                End If
                
                'Check Global Product key 
                sSubKeyNamePatches = ""
                sSubKeyNamePatches = GetRegConfigKey(sProductCode, iContext, sSid, True) & "Patches\" & sPatch
                If Not RegKeyExists(hDefKeyGlobal, sSubKeyNamePatches) Then
                    bPatchProdGlobalBroken = True
                Else
                    If RegReadDWordValue(hDefKeyGlobal, sSubKeyNamePatches, "State", sValue) Then 
                        arrRegPatch(1, UBound(arrRegPatch, 2)) = CInt(sValue)
                    End If 'RegReadDWordValue
                End If 'RegKeyExists(hDefKeyGlobal, sSubKeyNamePatches)
            End If 'arrType
        Next 'iName
    End If 'RegEnumValues
    'sTmpPatches should be an empty string now
    sTmpPatches = Replace(sTmpPatches, Chr(34), "")
    If (Not sTmpPatches = "") OR (NOT fPatchesOk) Then arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) &  ERR_BADMSPMETADATA & CSV

    'Find patches in other states (than 'Applied')
    '---------------------------------------------
    'Get a list from the global patches key
    sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, True) & "Patches\"
    bPatchGlobalMultiSzBroken = NOT RegReadValue(hDefKeyGlobal, sSubKeyName, "AllPatches", sAllPatches, REG_MULTI_SZ)
    
    'Sanity Check on metadata integrity
    If RegEnumKey(hDefKeyGlobal, sSubKeyName, arrKeys) AND CheckArray(arrKeys) Then
        bPatchGlobalOk = True
        sTmpPatches = sAllPatches
        For iKey = 0 To UBound(arrKeys)
            If InStr(sTmpPatches, arrKeys(iKey)) > 0 Then
                sTmpPatches = Replace(sTmpPatches, arrKeys(iKey), "")
            End If 'InStr
        Next 'iKey
        sTmpPatches = Replace(sTmpPatches, Chr(34), "")
        
        'Add the patches to the array
        For iKey = 0 To UBound(arrKeys)
            If Not InStr(sPatches, arrKeys(iKey)) > 0 Then
                ReDim Preserve arrRegPatch(REGDIM1, UBound(arrRegPatch, 2) + 1)
                sValue = ""
                If RegReadDWordValue(hDefKeyGlobal, sSubKeyName&arrKeys(iKey), "State", sValue) Then iState = CInt(sValue) Else iState = -1
                arrRegPatch(0, UBound(arrRegPatch, 2)) = GetExpandedGuid(arrKeys(iKey))
                arrRegPatch(1, UBound(arrRegPatch, 2)) = iState
            End If 'Not InStr
        Next 'iKey
    End If 'RegEnumKey
    
    iMspMax = 0 : If CheckArray(arrRegPatch) Then iMspMax = UBound(arrRegPatch, 2)
    If iMspMax > UBound(arrPatch, 3) Then ReDim Preserve arrPatch(UBound(arrMaster) , PATCH_COLUMNCOUNT, iMspMax)
    'Find the patch details
    ReadRegPatchDetails arrRegPatch, iPosMaster
End Sub 'FindRegPatches
'=======================================================================================================

Sub ReadPatchDetails(Patches, iPosMaster)
    Dim sTmpState
    On Error Resume Next
    
    Dim Patch, siSumInfo
    Dim iMspCnt
    
    iMspCnt = 0
    For Each Patch in Patches
        'Add patch data to array
        arrPatch(iPosMaster, PATCH_PRODUCT, iMspCnt) = arrMaster(iPosMaster, COL_PRODUCTCODE)
        arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) = Patch.PatchCode
        arrPatch(iPosMaster, PATCH_CSP, iMspCnt) = True
        arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) = Patch.PatchProperty("LocalPackage")
        arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) = False
        arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) = oFso.FileExists(arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt))
        arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt) = Patch.PatchProperty("DisplayName")
        If IsEmpty(arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt)) OR (arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt) =vbNullString) Then
            arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt) = GetRegMspProperty(GetCompressedGuid(Patch.PatchCode), iMspCnt, iPosMaster, arrMaster(iPosMaster, COL_CONTEXT), arrMaster(iPosMaster, COL_USERSID), "DisplayName")
        End If 'IsEmpty(iPosMaster, PATCH_DISPLAYNAME, iMspCnt)
        
        'InstallState:
        Select Case (Patch.State)
	        Case MSIPATCHSTATE_APPLIED, MSIPATCHSTATE_SUPERSEDED
	            If Patch.State = MSIPATCHSTATE_APPLIED Then sTmpState = "Applied" Else sTmpState = "Superseded"
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = sTmpState
	            If Patch.PatchProperty("Uninstallable") = "1" Then
		            arrPatch(iPosMaster, PATCH_UNINSTALLABLE, iMspCnt) = "Yes"
	            Else
		            arrPatch(iPosMaster, PATCH_UNINSTALLABLE, iMspCnt) = "No"
	            End If
	            If Not arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) Then 
	                arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & _
                     ERR_CATEGORYERROR & "Local patch package '" & arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) & _
	                 "' missing for patch '." & arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) & "'" & CSV
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Local patch package '" & arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) & _
	                 "' missing for patch '" & arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) & "'"
                End If
			    arrPatch(iPosMaster, PATCH_INSTALLDATE, iMspCnt) = Patch.PatchProperty("InstallDate")
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case MSIPATCHSTATE_OBSOLETED
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Obsoleted"
			    arrPatch(iPosMaster, PATCH_INSTALLDATE, iMspCnt) = Patch.PatchProperty("InstallDate")
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case MSIPATCHSTATE_REGISTERED
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Registered"
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = Patch.PatchProperty("MoreInfoURL")

	        Case Else
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Unknown Patchstate"
        End Select
        arrPatch(iPosMaster, PATCH_TRANSFORM, iMspCnt) = GetRegMspProperty(GetCompressedGuid(Patch.PatchCode), iMspCnt, iPosMaster, arrMaster(iPosMaster, COL_CONTEXT), arrMaster(iPosMaster, COL_USERSID), "PatchTransform")
        iMspCnt = iMspCnt + 1
    Next 'Patch
End Sub 'ReadPatchDetails
'=======================================================================================================

Sub ReadRegPatchDetails(arrRegPatch, iPosMaster)
    Dim hDefKey
    Dim sSid, sSubKeyName, sSubKeyNamePatch, sValue, sProductCode, sPatchCodeCompressed, sTmpState
    Dim iContext
    On Error Resume Next
    
    Dim Patch, siSumInfo
    Dim iMspCnt
    
    iContext = arrMaster(iPosMaster, COL_CONTEXT)
    sSid = arrMaster(iPosMaster, COL_USERSID)
    sProductCode = arrMaster(iPosMaster, COL_PRODUCTCODE)
    hDefKey = GetRegHive(iContext, sSid, True)
    sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, True) & "Patches\"
    For iMspCnt = 0 To UBound(arrRegPatch, 2)
        sPatchCodeCompressed = GetCompressedGuid(arrRegPatch(0, iMspCnt))
        'Add patch data to array
        arrPatch(iPosMaster, PATCH_PRODUCT, iMspCnt) = arrMaster(iPosMaster, COL_PRODUCTCODE)
        arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) = arrRegPatch(0, iMspCnt)
        arrPatch(iPosMaster, PATCH_CSP, iMspCnt) = True
        arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "LocalPackage")
        arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) = oFso.FileExists(arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt))
        arrPatch(iPosMaster, PATCH_DISPLAYNAME, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "DisplayName")
        Select Case (arrRegPatch(1, iMspCnt))
	        Case MSIPATCHSTATE_APPLIED, MSIPATCHSTATE_SUPERSEDED
	            If arrRegPatch(1, iMspCnt) = MSIPATCHSTATE_APPLIED Then sTmpState = "Applied" Else sTmpState = "Superseded"
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = sTmpState
	            arrPatch(iPosMaster, PATCH_UNINSTALLABLE, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "Uninstallable")
	            If Not arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) Then 
	                arrMaster(iPosMaster, COL_ERROR) = arrMaster(iPosMaster, COL_ERROR) & _
                     ERR_CATEGORYERROR & "Local patch package '" & arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) & _
	                 "' missing for patch '." & arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) & "'" & CSV
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "Local patch package '" & arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt) & _
	                 "' missing for patch '" & arrPatch(iPosMaster, PATCH_PATCHCODE, iMspCnt) & "'" 
                End If
			    arrPatch(iPosMaster, PATCH_INSTALLDATE, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "InstallDate")
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "MoreInfoURL")
	        Case MSIPATCHSTATE_OBSOLETED
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Obsoleted"
			    arrPatch(iPosMaster, PATCH_INSTALLDATE, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "InstallDate")
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "MoreInfoURL")
	        Case MSIPATCHSTATE_REGISTERED
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Registered"
			    arrPatch(iPosMaster, PATCH_MOREINFOURL, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "MoreInfoURL")
	        Case Else
	            arrPatch(iPosMaster, PATCH_PATCHSTATE, iMspCnt) = "Unknown Patchstate"
        End Select
        arrPatch(iPosMaster, PATCH_TRANSFORM, iMspCnt) = GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, "PatchTransform")
    Next 'iMspCnt
End Sub 'ReadRegPatchDetails
'=======================================================================================================

'Get a patch property without MSI API
Function GetRegMspProperty(sPatchCodeCompressed, iMspCnt, iPosMaster, iContext, sSid, sProperty)
    Dim hDefKey
    Dim sProductCode, sSubKeyName, sValue
    Dim siSumInfo
    On Error Resume Next
    
    GetRegMspProperty = ""
    sProductCode = arrMaster(iPosMaster, COL_PRODUCTCODE)
    hDefKey = GetRegHive(iContext, sSid, True)
    sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, True) & "Patches\" & sPatchCodeCompressed & "\"
    sValue = ""
    Select Case UCase(sProperty)
    Case "DISPLAYNAME"
        If RegReadStringValue(hDefKey, sSubKeyName, "DisplayName", sValue) Then GetRegMspProperty = sValue
        If GetRegMspProperty= "" Then
            If arrPatch(iPosMaster, PATCH_CPOK, iMspCnt) Then
                Err.Clear
                Set siSumInfo = oMsi.SummaryInformation(arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iMspCnt), MSIOPENDATABASEMODE_READONLY)
                If Not Err = 0 Then
                    GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
                Else
                    GetRegMspProperty = siSumInfo.Property(PID_TITLE)
                    If GetRegMspProperty = vbNullString Then GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
                End If 'Err = 0
            Else
                GetRegMspProperty = GetExpandedGuid(sPatchCodeCompressed)
            End If 'arrPatch(iPosMaster, PATCH_CPOK, iMspCnt)
        End If 'GetRegMspProperty= ""
    Case "INSTALLDATE"
        If RegReadStringValue(hDefKey, sSubKeyName, "Installed", sValue) Then GetRegMspProperty = sValue
    Case "LOCALPACKAGE"
        If iContext = MSIINSTALLCONTEXT_USERMANAGED Then iContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sSubKeyName = GetRegConfigPatchesKey(iContext, sSid, True) & sPatchCodeCompressed & "\"
        If RegReadStringValue(hDefKey, sSubKeyName, "LocalPackage", sValue) Then GetRegMspProperty = sValue
    Case "MOREINFOURL"
        If RegReadStringValue(hDefKey, sSubKeyName, "MoreInfoURL", sValue) Then GetRegMspProperty = sValue
    Case "PATCHTRANSFORM"
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey(sProductCode, iContext, sSid, False) & "Patches\" 
        If RegReadStringValue(hDefKey, sSubKeyName, sPatchCodeCompressed, sValue) Then GetRegMspProperty = sValue
    Case "UNINSTALLABLE"
        GetRegMspProperty = "No"
        If (RegReadDWordValue(hDefKey, sSubKeyName, "Uninstallable", sValue) AND sValue = "1") Then GetRegMspProperty = "Yes"
    Case Else
    End Select
End Function 'GetRegMspDisplayName
'=======================================================================================================

'Check if usage of PatchesEx API calls are safe for use.
'PatchesEx requires WI 3.x or higher
'If PatchesEx API calls fail this triggers a fallback to direct registry detection mode
Sub CheckPatchApi
    Dim sActiveSub, sErrHnd 
    sActiveSub = "CheckPatchApi" : sErrHnd = "" 
    On Error Resume Next
    Dim MspCheckEx, MspCheck
        
    'Defaults
    fPatchesExOk = False
    fPatchesOk = False
    iPatchesExError = 0
    'WI 2.x specific check
    Set MspCheck = oMsi.Patches(arrMaster(0, COL_PRODUCTCODE))
    If Err = 0 Then fPatchesOk = True Else Err.Clear
    'Ensure WI 3.x for PatchesEx checks
    If iWiVersionMajor = 2 Then 
        Exit Sub
    End If 'iWiVersionMajor = 2
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY, USERSID_EVERYONE, MSIINSTALLCONTEXT_ALL, MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = 8
        CheckError sActiveSub, sErrHnd 
    Else
        fPatchesExOk = True
        Exit Sub
    End If 'Not Err <> 0
        
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY, USERSID_EVERYONE, MSIINSTALLCONTEXT_USERMANAGED, MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_USERMANAGED
        CheckError sActiveSub, sErrHnd 
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_USERMANAGED
    
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY, USERSID_EVERYONE, MSIINSTALLCONTEXT_USERUNMANAGED, MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_USERUNMANAGED
        CheckError sActiveSub, sErrHnd 
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_USERUNMANAGED
    
    Set MspCheckEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY, MACHINESID, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_ALL)
    If Err <> 0 Then
        iPatchesExError = iPatchesExError + MSIINSTALLCONTEXT_MACHINE
        CheckError sActiveSub, sErrHnd
    End If 'Not Err <> 0 - MSIINSTALLCONTEXT_MACHINE
    
    Select Case iPatchesExError
        Case 8  CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL" & DOT
        Case 9  CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED" & DOT
        Case 10 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERUNMANAGED" & DOT
        Case 11 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_USERUNMANAGED" & DOT
        Case 12 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 13 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 14 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERUNMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case 15 CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PATCHESEX & "MSIINSTALLCONTEXT_ALL & MSIINSTALLCONTEXT_USERMANAGED & MSIINSTALLCONTEXT_USERUNMANAGED & MSIINSTALLCONTEXT_MACHINE" & DOT
        Case Else 
    End Select
End Sub 'CheckPatchApi
'=======================================================================================================

'Gather details from the cached patch package
Sub AddDetailsFromPackage
    On Error Resume Next

    Dim Element, Elements, Key, XmlDoc, Msp, SumInfo, Record
    Dim dicMspSequence, dicMspTmp
    Dim sFamily, sSeq, sMsp, sBaseSeqShort
    Dim iPosMaster, iPosPatch, iCnt, iIndex, iVersionMajor
    Dim fAdd
    Dim qView
    Dim arrTitle, arrTmpFam, arrTmpSeq, arrVer, arrPatchTargets
        
    'Defaults
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    Set dicMspSequence = CreateObject("Scripting.Dictionary")
    Set dicMspTmp = CreateObject("Scripting.Dictionary")
    'Initialize the patch index dictionary to assign each distinct PatchCode an index number
    iCnt = -1
    For iPosMaster = 0 To UBound(arrMaster)
        For iPosPatch = 0 to UBound(arrPatch, 3)
            If Not (IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))) AND (arrPatch(iPosMaster, PATCH_CSP, iPosPatch)) Then
                If NOT dicMspIndex.Exists(arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                    iCnt = iCnt + 1
                    dicMspIndex.Add arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch), iCnt
                    dicMspTmp.Add arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch), arrPatch(iPosMaster, PATCH_LOCALPACKAGE, iPosPatch)
                End If
            End If
        Next 'iPosPatch
    Next 'iPosMaster
    'Initialize the PatchFiles array
    ReDim arrMspFiles(dicMspIndex.Count - 1, MSPFILES_COLUMNCOUNT - 1)
    'Open the .msp for reading to add the data to the array
    For Each Key in dicMspTmp.Keys
        iPosPatch = dicMspIndex.Item(Key)
        sMsp = dicMspTmp.Item(Key)
        Err.Clear
        Set Msp = oMsi.OpenDatabase (sMsp, MSIOPENDATABASEMODE_PATCHFILE)
        Set SumInfo = Msp.SummaryInformation
        If Not Err = 0 Then
            Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_MSPOPENFAILED & sMsp
        Else
            arrMspFiles(iPosPatch, MSPFILES_LOCALPACKAGE) = sMsp
            arrMspFiles(iPosPatch, MSPFILES_PATCHCODE) = Key
            arrMspFiles(iPosPatch, MSPFILES_TARGETS) = SumInfo.Property(PID_TEMPLATE)
            arrMspFiles(iPosPatch, MSPFILES_TABLES) = GetPatchTables(Msp)
            If InStr(arrMspFiles(iPosPatch, MSPFILES_TABLES), "MsiPatchMetadata") > 0 Then
                Set qView = Msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='KBArticle Number'")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    arrMspFiles(iPosPatch, MSPFILES_KB) = UCase(Record.StringData(2))
                    arrMspFiles(iPosPatch, MSPFILES_KB) = Replace(arrMspFiles(iPosPatch, MSPFILES_KB), "KB","")
                Else
                    arrMspFiles(iPosPatch, MSPFILES_KB) = ""
                End If
                qView.Close
                Set qView = Msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = Record.StringData(2)
                Else
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = ""
                End If
                qView.Close
            Else
                arrMspFiles(iPosPatch, MSPFILES_KB) = ""
                arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = ""
            End If
            arrTitle = Split(SumInfo.Property(PID_TITLE), ";")
            If arrMspFiles(iPosPatch, MSPFILES_KB) = "" Then
                If UBound(arrTitle) > 0 Then
                    arrMspFiles(iPosPatch, MSPFILES_KB) = arrTitle(1)
                End If
            End If
            If arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = "" Then
                If UBound(arrTitle) > 0 Then
                    arrMspFiles(iPosPatch, MSPFILES_PACKAGE) = arrTitle(1)
                End If
            End If
            If InStr(arrMspFiles(iPosPatch, MSPFILES_TABLES), "MsiPatchSequence") > 0 Then
                Set qView = Msp.OpenView("SELECT `PatchFamily`, `Sequence`, `Attributes` FROM MsiPatchSequence")
                qView.Execute : Set Record = qView.Fetch()
                If Not Record Is Nothing Then
                    Do Until Record Is Nothing
                        arrMspFiles(iPosPatch, MSPFILES_FAMILY) = arrMspFiles(iPosPatch, MSPFILES_FAMILY) & LCase(Record.StringData(1)) & ","
                        sSeq = Record.StringData(2)
                        If NOT InStr(sSeq, ".") > 0 Then sSeq = GetVersionMajor(Left(arrMspFiles(iPosPatch, MSPFILES_TARGETS), 38)) & ".0." & sSeq & ".0"
                        arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) & sSeq & ","
                        arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = Record.StringData(3)
                        Set Record = qView.Fetch()
                    Loop
                    arrMspFiles(iPosPatch, MSPFILES_FAMILY) = RTrimComma(arrMspFiles(iPosPatch, MSPFILES_FAMILY))
                    arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = RTrimComma(arrMspFiles(iPosPatch, MSPFILES_SEQUENCE))
                Else
                    arrMspFiles(iPosPatch, MSPFILES_FAMILY) = ""
                    arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = "0"
                    arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = "0"
                End If
                qView.Close
            Else
                arrMspFiles(iPosPatch, MSPFILES_FAMILY) = ""
                arrMspFiles(iPosPatch, MSPFILES_SEQUENCE) = "0"
                arrMspFiles(iPosPatch, MSPFILES_ATTRIBUTE) = "0"
            End If
        End If
    Next 'Key

    For iPosMaster = 0 To UBound(arrMaster)
        dicMspSequence.RemoveAll
        sBaseSeqShort = ""
        arrVer = Split(arrMaster(iPosMaster, COL_PRODUCTVERSION), ".")
        If UBound(arrVer)>1 Then sBaseSeqShort = arrVer(2)
        For iPosPatch = 0 to UBound(arrPatch, 3)
            If Not (IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))) AND (arrPatch(iPosMaster, PATCH_CSP, iPosPatch)) Then
                iIndex = -1
                iIndex = dicMspIndex.Item(arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))
                If NOT iIndex = -1 Then
                    arrPatch(iPosMaster, PATCH_KB, iPosPatch) = arrMspFiles(iIndex, MSPFILES_KB)
                    arrPatch(iPosMaster, PATCH_PACKAGE, iPosPatch) = arrMspFiles(iIndex, MSPFILES_PACKAGE)
                    Set arrTmpSeq = Nothing
                    arrTmpSeq = Split(arrMspFiles(iIndex, MSPFILES_SEQUENCE), ",")
                    arrPatch(iPosMaster, PATCH_SEQUENCE, iPosPatch) = arrTmpSeq(0)
                    Set arrTmpFam = Nothing
                    arrTmpFam = Split(arrMspFiles(iIndex, MSPFILES_FAMILY), ",")
                    If arrMspFiles(iIndex, MSPFILES_ATTRIBUTE) = "1" Then
                        iCnt = -1
                        For iCnt = 0 To UBound(arrTmpSeq)
                            sFamily= "" : sSeq= ""
                            sFamily = arrTmpFam(iCnt)
                            sSeq = arrTmpSeq(iCnt)
                            fAdd = False
                            Select Case Len(sSeq)
                            Case 4
                                fAdd = (sSeq > sBaseSeqShort)
                            Case Else
                                fAdd = (sSeq > arrMaster(iPosMaster, COL_PRODUCTVERSION))
                            End Select
                            If fAdd Then 
                                If dicMspSequence.Exists(sFamily) Then
                                    If (sSeq > dicMspSequence.Item(sFamily)) Then dicMspSequence.Item(sFamily) =sSeq
                                Else
                                    dicMspSequence.Add sFamily, sSeq
                                End If
                            End If 'fAdd
                        Next 'iCnt
                    End If 'Attributes = 1
                    iCnt = -1
                    For iCnt = 0 To UBound(arrTmpSeq)
                        sFamily= "" : sSeq= ""
                        sFamily = arrTmpFam(iCnt)
                        sSeq = arrTmpSeq(iCnt)
                        If dicPatchLevel.Exists(sFamily) Then
                            If (sSeq > dicPatchLevel.Item(sFamily)) Then dicPatchLevel.Item(sFamily) = sSeq
                        Else
                            dicPatchLevel.Add sFamily, sSeq
                        End If
                    Next 'iCnt

                End If 'iIndex
            End If
        Next 'iPosPatch
        For Each Key in dicMspSequence.Keys
            arrMaster(iPosMaster, COL_PATCHFAMILY) =arrMaster(iPosMaster, COL_PATCHFAMILY)&Key& ":"&dicMspSequence.Item(Key)& ","
        Next 'Key
        RTrimComma arrMaster(iPosMaster, COL_PATCHFAMILY)
    Next 'iPosMaster
End Sub 'AddDetailsFromPackage
'=======================================================================================================

'Return the tables of a given .msp file
Function GetPatchTables(Msp)

Dim ViewTables, Table
Dim sTables

    On Error Resume Next
    sTables = ""
    Set Table = Nothing
    Set ViewTables = Msp.OpenView("SELECT `Name` FROM `_Tables` ORDER BY `Name`")
    ViewTables.Execute
    Do
        Set Table = ViewTables.Fetch
        If Table Is Nothing then Exit Do
        sTables = sTables&Table.StringData(1)& ","
    Loop
    ViewTables.Close
    GetPatchTables=RTrimComma(sTables)
End Function 'GetPatchTables

'=======================================================================================================
'Module ARP - Add Remove Programs
'=======================================================================================================
'Build array with data from Add/Remove Programs.
'This is essential for Office 2007 family applications 
Sub ARPData
    On Error Resume Next

    'Filter for >O12 products since the logic is proprietary to >O12
     FindArpOMsiConfigProductsEntry 
     AddArpOConfigProductsToMaster 
     AddSystemComponentFlagToMaster 
     ValidateSystemComponentProducts 
     FindArpConfigMsi 
End Sub 'ARPData
'=======================================================================================================

'Identify the core configuration .msi for multi .msi ARP entries
Sub FindArpConfigMsi
    Dim n, i, k
    Dim hDefKey
    Dim sSubKeyName, sCurKey, sName, sValue, sArpProd, sArpProdWW
    Dim arrKeys, arrValues
    Dim ProdId
    On Error Resume Next

    hDefKey = HKEY_LOCAL_MACHINE
    If Not CheckArray(arrArpProducts) Then 
        If Not UBound(arrArpProducts) = -1 Then 
            Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_NOARRAY & " Terminated ARP detection"
        End If
        Err.Clear
        Exit Sub
    End If
   
    For n = 0 To UBound(arrArpProducts)
        k = -1
        sArpProd = arrArpProducts(n, COL_CONFIGNAME)
        If Not InStr(sArpProd, ".") > 0 Then sArpProdWW = sArpProd & "WW"
        sName = "PackageIds"
        If UCase(Left(sArpProd, 9)) = "OFFICE14." OR UCase(Left(sArpProd, 9)) = "OFFICE15." OR UCase(Left(sArpProd, 9)) = "OFFICE16." Then 
            sName = "PackageRefs"
            sArpProdWW = Right(sArpProd, Len(sArpProd) - 9) & "WW"
        End If
        sSubKeyName = REG_ARP & arrArpProducts(n, COL_CONFIGNAME) & "\"

        If RegReadMultiStringValue(hDefKey, sSubKeyName, sName, arrValues) Then
            i = 0
            'The target product is usually at the last entry of the array
            'Check last entry first. If this fails search the array
            If InStr(UCase(arrValues(UBound(arrValues))), UCase(sArpProdWW)) > 0 Then
                k = UBound(arrValues)
            ElseIf k = -1 AND sName = "PackageRefs" Then
                ' trim away the OFFIC1X. part from the string
                If InStr(UCase(arrValues(UBound(arrValues))), UCase(Right(sArpProd, Len(sArpProd) - 9))) > 0 Then
                    k = UBound(arrValues)
                End If
            ElseIf k = -1 AND InStr(UCase(arrValues(UBound(arrValues))), UCase(sArpProd)) > 0 Then
                k = UBound(arrValues)
            Else
                For Each ProdId in arrValues
                    If InStr(UCase(ProdID), UCase(sArpProdWW)) > 0 Then 
                        k = i
                        Exit For
                    End If 'InStr
                    i = i + 1
                Next 'ProdId
                'try without the 'WW' extension
                If k = -1 Then
                    For Each ProdId in arrValues
                        If InStr(UCase(ProdID), UCase(sArpProd)) > 0 Then 
                            k = i
                            Exit For
                        End If 'InStr
                        i = i + 1
                    Next 'ProdId
                End If 'k = -1
            End If 'InStr
            If k = -1 Then
                k = UBound(arrValues)
            End If
        Else
        End If 'RegReadMultiStringValue
        
        If Not k = -1 Then 
            'found a matching 'Config Guid' for the product
            arrArpProducts(n, COL_CONFIGPACKAGEID) = arrValues(k) 
        End If 'k
    Next 'n
End Sub 'FindArpConfigMsi
'=======================================================================================================

'Check if every Office product flagged as 'SystemComponent = 1' has a parent registered
Sub ValidateSystemComponentProducts
    Dim n
    On Error Resume Next
    
    For n = 0 To UBound (arrMaster)
        If Not Err = 0 Then Exit For
        If arrMaster(n, COL_ISOFFICEPRODUCT) Then
            If (NOT arrMaster(n, COL_SYSTEMCOMPONENT) = 0)  AND (arrMaster(n, COL_ARPPARENTCOUNT) = 0) Then
                Select Case Mid(arrMaster(n, COL_PRODUCTCODE), 11, 4)
                Case "0010","00B9","008C","008F","007E","00DD"
                    ' don't add warning
                Case Else
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYNOTE, ERR_ORPHANEDITEM & DSV & arrMaster(n, COL_PRODUCTNAME) & DSV & arrMaster(n, COL_PRODUCTCODE)
                    arrMaster(n, COL_NOTES) = arrMaster(n, COL_NOTES) & ERR_CATEGORYNOTE & ERR_ORPHANEDITEM & CSV
                End Select
            End If 'Not arrMaster
        End If 'IsOfficeProduct
    Next 'n
End Sub
'=======================================================================================================

'Adds the ARP config product information to the individual products
Sub AddArpOConfigProductsToMaster
    Dim i, n
    On Error Resume Next

    For n = 0 To UBound(arrMaster)
        If Not Err = 0 Then Exit For
        If arrMaster(n, COL_ISOFFICEPRODUCT) Then
            For i = 0 To UBound(arrArpProducts)
                If InStr(arrArpProducts(i, COL_ARPALLPRODUCTS), arrMaster(n, COL_PRODUCTCODE)) > 0 Then
                    arrMaster(n, COL_ARPPARENTCOUNT) = arrMaster(n, COL_ARPPARENTCOUNT) + 1
                    arrMaster(n, COL_ARPPARENTS) = arrMaster(n, COL_ARPPARENTS) & "," & arrArpProducts(i, COL_CONFIGNAME)
                End If
            Next 'i
        End If
    Next 'n
End Sub
'=======================================================================================================

'Add applications 'SystemComponent' flag from ARP to Master array
Sub AddSystemComponentFlagToMaster
    Dim sValue, sSubKeyName
    Dim n
    On Error Resume Next

    For n = 0 To UBound(arrMaster)
        If Not Err = 0 Then Exit For
        sSubKeyName = REG_ARP & arrMaster(n, COL_PRODUCTCODE) & "\"
        If RegReadDWordValue(HKEY_LOCAL_MACHINE, sSubKeyName, "SystemComponent", sValue) Then
            arrMaster(n, COL_SYSTEMCOMPONENT) = sValue
        Else
            arrMaster(n, COL_SYSTEMCOMPONENT) = 0
        End If 'RegReadDWordValue "SystemComponent"
    Next 'n
End Sub
'=======================================================================================================

'-------------------------------------------------------------------------------
'   FindArpOMsiConfigProductsEntry
'
'   Find parent/configuration entries from 'Add/Remove Programs' for 
'   multi MSI Office products
'   ARP entries flagged as Systemcomponent will not display in ARP
'   From Office 2007 on ARP groups the multi MSI configuration together.
'-------------------------------------------------------------------------------
Sub FindArpOMsiConfigProductsEntry
    Dim ArpProd, ProductCode, Key, dicKey
    Dim n, i, j, iMaxVal, iArpCnt, iPosMaster, iLoop
    Dim fNoSysComponent, fOConfigEntry, fFoundConfigProductCode, fUninstallString
    Dim hDefKey
    Dim sSubKeyName, sCurKey, sName, sValue, sNames, sArpProductName, sRegArp, sMondo
    Dim arrKeys, arrValues, arrNames, arrTypes
    Dim dicArpTmp
    Dim tModulStart, tModulEnd
    On Error Resume Next
    
    iLoop = 0 : iMaxVal = 0: iArpCnt = 0
    Set dicArpTmp = CreateObject("Scripting.Dictionary")
    
    hDefKey = HKEY_LOCAL_MACHINE
    sRegArp = REG_ARP
    sSubKeyName = sRegArp
    Set dicMissingChild = CreateObject("Scripting.Dictionary")
    For Each ArpProd in dicArp.Keys
        'Read the contents of the key
        sCurKey = dicArp.Item(ArpProd) & "\"
        fNoSysComponent = True
        fOConfigEntry = False
        sNames = "" : Redim arrNames(-1) : Redim arrTypes(-1)
        If RegEnumValues(hDefKey, sCurKey, arrNames, arrTypes) Then sNames = Join(arrNames)
        ' Find/count multi MSI Office ARP entries
        FindArpOMultiMsiConfigProducts ArpProd, iArpCnt, sCurKey, sNames, fOConfigEntry, arrNames, dicArpTmp
    Next 'ArpProd

    ' initialize arrArpProducts
    If iArpCnt > 0 Then 
        iArpCnt = iArpCnt - 1
        Redim arrArpProducts(iArpCnt, ARP_CHILDOFFSET)
    Else
        Redim arrArpProducts(-1)
    End If
    ' copy the identified configuration parents to the actual target array and add additional fields
    n = 0
    For Each dicKey in dicArpTmp.Keys
        ArpProd = dicArpTmp.Item(dicKey)
        sName = "ProductCodes"
        sSubKeyName = dicKey
        If RegReadMultiStringValue(hDefKey, sSubKeyName, sName, arrValues) Then
            AddArpOMultiMsiProperties n, iMaxVal, iArpCnt, ArpProd, sSubKeyName, arrValues, arrArpProducts
        Else
            AddArpC2RProperties n, iMaxVal, iArpCnt, ArpProd, arrArpProducts
        End If 
        n = n + 1
    Next 'ArpProd

End Sub 'FindArpOMsiConfigProductsEntry

'-------------------------------------------------------------------------------
'   FindArpOMultiMsiConfigProducts
'
'   From Office 2007 on ARP groups the multi MSI configuration together.
'-------------------------------------------------------------------------------
Sub FindArpOMultiMsiConfigProducts (ArpProd, iArpCnt, sCurKey, sNames, fOConfigEntry, arrNames, dicArpTmp)
    Dim sName, sValue
    On Error Resume Next
    
    If InStr(sNames, "PackageIds") OR _
        InStr(sNames, "PackageRefs") OR _
        UCase(Left(ArpProd, 9)) = "OFFICE14." OR _
        UCase(Left(ArpProd, 9)) = "OFFICE15." OR _
        UCase(Left(ArpProd, 9)) = "OFFICE16." Then fOConfigEntry = True

    If InStr(LCase(sCurKey), LCase("\ClickToRun\REGISTRY\MACHINE\Software")) > 0 Then fOConfigEntry = False
            
    If fOConfigEntry Then 
        sName = "SystemComponent"
        'If RegDWORD 'SystemComponent' = 1 then it's not showing up in ARP
        If RegReadDWordValue(HKLM, sCurKey, sName, sValue) Then
            If CInt(sValue) = 1 Then
                CacheLog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYWARN, "Office configuration entry that has been hidden from ARP: " & ArpProd
            End If
        End If 'RegReadValue "SystemComponent"
        If fOConfigEntry Then
            ' found ARP Parent entry
            ' only care if it's not a 'helper' entry 
            If UBound(arrNames) > 5 OR InStr(UCase(ArpProd), ".MONDO") > 0 Then
                If NOT dicArpTmp.Exists(sCurKey) Then dicArpTmp.Add sCurKey, ArpProd
                iArpCnt = iArpCnt + 1
                Cachelog LOGPOS_RAW, LOGHEADING_NONE, Null, "ARP Entry: " & ArpProd
            Else
                Cachelog LOGPOS_RAW, LOGHEADING_NONE, Null, "Bad ARP Entry: " & ArpProd
                Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_OFFSCRUB_TERMINATED
                Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_BADARPMETADATA & GetHiveString(hDefKey) & "\" & sCurKey
            End If
        End If
    End If
End Sub 'FindArpOMultiMsiConfigProducts


'-------------------------------------------------------------------------------
'   AddArpOMultiMsiProperties
'
'   Add details for Multi MSI products
'-------------------------------------------------------------------------------
Sub AddArpOMultiMsiProperties (n, iMaxVal, iArpCnt, ArpProd, sSubKeyName, arrValues, arrArpProducts)
    Dim i, j, iPosMaster
    Dim sValue, sArpProductName, sConfigMsiId, sInstallSource
    Dim ProductCode, ArpKey
    Dim fFoundConfigProductCode
    On Error Resume Next

    ' fill the array set
    arrArpProducts(n, COL_CONFIGNAME) = ArpProd
    arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "MSI"
    RegReadStringValue HKLM, sSubKeyName, "DisplayVersion", sValue
    arrArpProducts(n, ARP_PRODUCTVERSION) = sValue
    If UBound (arrValues) + ARP_CHILDOFFSET > iMaxVal Then 
        iMaxVal = UBound (arrValues) + ARP_CHILDOFFSET
        Redim Preserve arrArpProducts(iArpCnt, iMaxVal)
    End If 'UBound(arrValues)
    j = ARP_CHILDOFFSET
    For Each ProductCode in arrValues
        arrArpProducts(n, COL_ARPALLPRODUCTS) = arrArpProducts(n, COL_ARPALLPRODUCTS) & ProductCode
        arrArpProducts(n, j) = ProductCode
        ' validate that the product is installed
        If NOT dicProducts.Exists(ProductCode) Then
            Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_MISSINGCHILD & "Config Product: " & ArpProd & CSV & "Missing Product: " & ProductCode
            If NOT dicMissingChild.Exists(ProductCode) Then dicMissingChild.Add ProductCode, ArpProd
        End If
        j = j + 1
    Next 'ProductCode
            
    ' identify the config ProductCode
    ' first try to match the full ARP Name by enum the linked ProductCodes
    fFoundConfigProductCode = False
    sArpProductName = ""
    sArpProductName = GetArpProductName(ArpProd)
    For Each ProductCode in arrValues
        iPosMaster = GetArrayPosition(arrMaster, ProductCode)
        If NOT iPosMaster = -1 Then
            If LCase(Trim(sArpProductName)) = LCase(Trim(arrMaster(GetArrayPosition(arrMaster, ProductCode), COL_PRODUCTNAME))) Then
                arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = ProductCode
                fFoundConfigProductCode = True
                Exit For
            End If
        End If
    Next 'ProductCode
    
    If NOT fFoundConfigProductCode Then
        If RegReadStringValue(HKLM, sSubKeyName, "UninstallString", sValue) Then
            sConfigMsiId = UCase(sValue)
            If InStr(sConfigMsiId, "/UNINSTALL ") > 0 Then
                sConfigMsiId = Mid(sConfigMsiId, InStr(sConfigMsiId, "/UNINSTALL ") + 10)
                sConfigMsiId = Left(sConfigMsiId, InStr(sConfigMsiId, " /DLL") - 1)
            End If
            For Each ProductCode in arrValues
                If dicArp.Exists(ProductCode) Then
                    If RegReadStringValue(HKLM, dicArp.Item(ProductCode), "InstallSource", sInstallSource) Then
                        If InStr(UCase(sInstallSource), sConfigMsiId) > 0 Then
                            iPosMaster = GetArrayPosition(arrMaster, ProductCode)
                            If NOT iPosMaster = -1 Then
                                arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = ProductCode
                                fFoundConfigProductCode = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
            
    If NOT fFoundConfigProductCode Then
        i = 0
        For Each ProductCode in arrValues
            If (NOT Mid(ProductCode, 11, 4) = "002A") AND (Mid(ProductCode, 16, 4) = "0000") _ 
                OR (Mid(ProductCode, 12, 2) = "10") _ 
            Then
                i = i + 1
                arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = ProductCode
            End If
        Next 'ProductCode
                
        ' ensure to not have ambigious hits
        If i > 1 Then
            Dim MsiDb, Record
            Dim qView
            For Each ProductCode in arrValues
                If (NOT Mid(ProductCode, 11, 4) = "002A") AND (Mid(ProductCode, 16, 4) = "0000") _ 
                    OR (Mid(ProductCode, 12 , 2) = "10") _ 
                Then
                    Set MsiDb = oMsi.OpenDatabase(arrMaster(GetArrayPosition(arrMaster, ProductCode), COL_CACHEDMSI), MSIOPENDATABASEMODE_READONLY)
	                Set qView = MsiDb.OpenView("SELECT * FROM Property WHERE `Property` = 'SetupExeArpId'")
	                qView.Execute 
	                Set Record = qView.Fetch()
	                If Not Record Is Nothing Then
	                    arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = ProductCode
	                    qView.Close
	                    Exit For
	                End If 'Is Nothing
                End If
            Next 'ProductCode
        End If
        if i = 0 Then
            arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = arrValues(UBound(arrValues))
        End If
    End If 'NOT fFoundConfigProductCode
End Sub 'AddArpOMultiMsiProperties

'-------------------------------------------------------------------------------
'   AddArpC2RProperties
'
'   Add details for C2R products
'-------------------------------------------------------------------------------
Sub AddArpC2RProperties (n, iMaxVal, iArpCnt, ArpProd, arrArpProducts)
    Dim Key, prod
    Dim sValue, sMondo
    Dim i, j, iPosMaster
    On Error Resume Next

    arrArpProducts(n, COL_CONFIGNAME) = ArpProd
    arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "C2R"
    If 1 + ARP_CHILDOFFSET > iMaxVal Then 
        iMaxVal = 1 + ARP_CHILDOFFSET
        Redim Preserve arrArpProducts(iArpCnt, iMaxVal)
    End If 
    
    ' V1
    For Each Key in dicArp.Keys
        If Len(Key) = 38 Then
            If Mid(Key, 11, 4) = "006D" Then
                arrArpProducts(n, COL_ARPALLPRODUCTS) = Key
                arrArpProducts(n, ARP_CHILDOFFSET) = Key
                arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = Key
            End If
        End If
    Next 'Key

    ' V2
    sMondo = ""
    If InStr(UCase(ArpProd), ".MONDO") > 0 Then 
        sMondo = Mid(ArpProd, 7, 2) & "0000-000F-0000-0000-0000000FF1CE}"
    Else
        RegReadValue HKLM, dicArp.Item(ArpProd), "UninstallString", sValue, "REG_SZ"
        If (InStr(LCase(sValue), "productstoremove=") > 0 AND InStr(sValue, ".15_") > 0 ) OR _
           (InStr(LCase(sValue), "productstoremove=") > 0 AND InStr(sValue, ".16_") > 0  ) Then
        	prod = Mid(sValue, InStr(sValue, "productstoremove="))
	        prod = Replace(prod, "productstoremove=","")
	        If InStr(prod, "_") > 0 Then
	            prod = Left(prod, InStr(prod, "_") - 1)
	            If InStr(prod, ".1") > 0 Then
	                prod = Left(prod, InStr(prod, ".1") - 1)
                    sValue = prod
	            End If
	        End If
        Else
            If InStr(LCase(sValue), "microsoft office 1") > 0 Then
                sValue = Mid(sValue, InStr(sValue, "Microsoft Office ") + 17, 2)
            End If
        End If
        sMondo = sValue & "0000-000F-0000-0000-0000000FF1CE}"
    End If

    iPosMaster = GetArrayPositionFromPattern(arrMaster, sMondo)
    If InStr(UCase(ArpProd), ".MONDO") > 0 Then
        arrArpProducts(n, COL_CONFIGINSTALLTYPE) = "VIRTUAL"
        If NOT iPosMaster = -1 Then
            arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = arrMaster(iPosMaster, COL_PRODUCTCODE)
        End If
        If UBound (arrMVProducts) + ARP_CHILDOFFSET > iMaxVal Then 
            iMaxVal = UBound (arrMVProducts) + ARP_CHILDOFFSET
            Redim Preserve arrArpProducts(iArpCnt, iMaxVal)
        End If 'UBound(arrMVProducts)
        ' fill the array set
        j = ARP_CHILDOFFSET
        i = 0
        For i = 0 To UBound(arrMVProducts)
            'If arrMaster(i, COL_VIRTUALIZED) = 1 Then
                arrArpProducts(n, COL_ARPALLPRODUCTS) = arrArpProducts(n, COL_ARPALLPRODUCTS) & arrMVProducts(i, COL_PRODUCTCODE)
                arrArpProducts(n, j) = arrMVProducts(i, COL_PRODUCTCODE)
                j = j + 1
            'End If
        Next 'i
    Else
        If NOT iPosMaster = -1 Then
            arrArpProducts(n, COL_ARPALLPRODUCTS) = arrMaster(iPosMaster, COL_PRODUCTCODE)
            arrArpProducts(n, ARP_CHILDOFFSET) = arrMaster(iPosMaster, COL_PRODUCTCODE)
            arrArpProducts(n, ARP_CONFIGPRODUCTCODE) = arrMaster(iPosMaster, COL_PRODUCTCODE)
        End If
    End If

End Sub 'AddArpC2RProperties

'=======================================================================================================
'Module (O)SPP - Licensing Data / LicenseState
'=======================================================================================================

'-------------------------------------------------------------------------------
'   OsppCollect
'
'   Use WMI to query the license details for installed Office products
'   The license details of the products are mapped to the OSPP data based on the 
'   "Configuration Productname"
'-------------------------------------------------------------------------------
Sub OsppCollect
    Dim iArpCnt, iPosMaster, iVPCnt, iVersionMajor
    Dim sText, sPrid
    Dim dicInstalledLics, dicInstalledActiveLics
    Dim arrPrids
    On Error Resume Next

    Set dicInstalledLics = CreateObject("Scripting.Dictionary")
    Set dicInstalledActiveLics = CreateObject("Scripting.Dictionary")

     GetSPPAllInstalledLics (dicInstalledLics)
    
     GetSPPActiveInstalledLics (dicInstalledActiveLics)

    If CheckArray(arrArpProducts) Then
        For iArpCnt = 0 To UBound(arrArpProducts)
            'Get the link to the master array
            iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, ARP_CONFIGPRODUCTCODE))
            If iPosMaster > -1 Then
                'OSPP is first used by O14
                iVersionMajor = CInt(GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)))
                If (arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)) AND iVersionMajor > 13 Then
                    'Get the Config Productname
                    sText = "" : sText = arrArpProducts(iArpCnt, COL_CONFIGNAME)
                    If InStr(sText, ".") > 0 Then sText = Mid(sText, InStr(sText, ".") + 1)
                    arrPrids = Split(GetProductReleaseIdFromPrimaryProductId(iVersionMajor, arrMaster(iPosMaster, COL_PRIMARYPRODUCTID)), ";")
                    If IsArray(arrPrids) Then
                        For Each sPrid in arrPrids
                            If InStr(UCase(sPrid), UCase(sText)) > 0 Then
                                OsppCollectForArray sPrid, arrMaster, dicInstalledLics, dicInstalledActiveLics, iVersionMajor, iPosMaster, COL_ACIDLIST, COL_LICENSELIST, COL_OSPPLICENSE, COL_OSPPLICENSEXML
                            End If
                        Next
                    End If
                End If 'VersionMajor > 14
            End If 'iPosMaster > -1
        Next 'iArpCnt
    End If 'arrArpProducts

    If UBound(arrVirt2Products) > -1 Then
        For iVPCnt = 0 To UBound(arrVirt2Products, 1)
            sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)
            OsppCollectForArray sText, arrVirt2Products, dicInstalledLics, dicInstalledActiveLics, 15, iVPCnt, VIRTPROD_ACIDLIST, VIRTPROD_LICENSELIST, VIRTPROD_OSPPLICENSE, VIRTPROD_OSPPLICENSEXML
        Next 'iVPCnt
    End If 'arrVirt2Products

    If UBound(arrVirt3Products) > -1 Then
        For iVPCnt = 0 To UBound(arrVirt3Products, 1)
            sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)
            OsppCollectForArray sText, arrVirt3Products, dicInstalledLics, dicInstalledActiveLics, 16, iVPCnt, VIRTPROD_ACIDLIST, VIRTPROD_LICENSELIST, VIRTPROD_OSPPLICENSE, VIRTPROD_OSPPLICENSEXML
        Next 'iVPCnt
    End If 'arrVirt2Products
End Sub 'OsppCollect

'-------------------------------------------------------------------------------
'   OsppCollectForArray
'
'   Handler for passed in array (ArrMaster, arrVirtX)
'-------------------------------------------------------------------------------
Sub OsppCollectForArray (sConfigName, arrProd, dicInstalledLics, dicInstalledActiveLics, iVersionMajor, iPos, iPosAcidList, iPosLicList, iPosOsppLic, iPosOsppLicXml)
    On Error Resume Next
    Dim sSkuList, sAcidList, sXmlLogLine, sOsppLicenses
    Dim key
    Dim dicPridAcid, dicThisProductAll, dicLicChannel, dicThisProductInUse, dicThisProductLicUnused

    Set dicThisProductAll = CreateObject("Scripting.Dictionary")
    Set dicPridAcid = CreateObject("Scripting.Dictionary")
    Set dicThisProductInUse = CreateObject("Scripting.Dictionary")
    Set dicThisProductLicUnused = CreateObject("Scripting.Dictionary")
    Set dicLicChannel = CreateObject("Scripting.Dictionary")

    sSkuList = ""
    sAcidList = ""
    sXmlLogLine = ""
    MapConfigNameToLicense sConfigName, iVersionMajor, sSkuList, sAcidList
    ArrProd(iPos, iPosAcidList) = sAcidList
    ArrProd(iPos, iPosLicList) = sSkuList

    GetPRIdAcidDic dicPridAcid, sConfigName, iVersionMajor

    sOsppLicenses = "Possible Licenses;"
    For Each key in dicInstalledLics.Keys
        If dicPridAcid.Exists(key) Then
            AddToDic dicThisProductAll, key, dicPridAcid.Item(key)
            sOsppLicenses = sOsppLicenses & key & " " & dicPridAcid.Item(key) & ","
        End If
    Next
    ArrProd(iPos, iPosOsppLic) = sOsppLicenses

    For Each key in dicInstalledActiveLics.Keys
        If dicPridAcid.Exists(key) Then AddToDic dicThisProductInUse, key, dicPridAcid.Item(key)
    Next
    For Each key in dicThisProductAll.Keys
        AddToDic dicLicChannel, key, GetLicTypeFromAcid(key)
    Next

    AddLicenseData dicThisProductInUse, dicLicChannel, ArrProd, iPos, iPosOsppLic, iPosOsppLicXml

    For Each key in dicThisProductAll
        If NOT dicThisProductInUse.Exists(key) Then
            dicThisProductLicUnused.RemoveAll
            dicThisProductLicUnused.Add "ID", key
            dicThisProductLicUnused.Add "Name", dicThisProductAll.Item(key)
            dicThisProductLicUnused.Add "Channel", dicLicChannel.Item(key)
            dicThisProductLicUnused.Add "LicenseStatus", 0
            dicThisProductLicUnused.Add "LicenseStatusString", "UNLICENSED"
            dicThisProductLicUnused.Add "LicenseErrorCode","0xC004F014"
            dicThisProductLicUnused.Add "LicenseErrorCodeExplain", GetLicErrDesc("C004F014")
            arrProd(iPos, iPosOsppLicXml) = GetLicXmlString(arrProd(iPos, iPosOsppLicXml), dicThisProductLicUnused)
        End If
    Next
End Sub 'OsppCollectForArray

'-------------------------------------------------------------------------------
'   AddLicenseData
'
'   Add license data for the passed in ACID
'-------------------------------------------------------------------------------
Sub AddLicenseData (dicProdLics, dicLicChannel, arrProd, iPos, iPosOsppLic, iPosOsppLicXml)
    On Error Resume Next
    Dim key, item
    Dim fActive

    For Each key in dicProdLics.Keys
        Select Case dicLicChannel.Item(key)
        Case "Subscription","ZeroGrace","Trial","SubTrial","SubTest","SubPrepid","Grace","DeltaTrial","BypassTrial365","BypassTrial180","Bypass30"
            AddSPPSubscriptionProps key, dicLicChannel.Item(key), arrProd, iPos, iPosOsppLic, iPosOsppLicXml
        Case "Retail","PrepidBypass","PIN","OEM_Perp","OEM_ARM","MAKC2R","MAK_AE","MAK","Bypass","BypassTrial"
            AddSPPPerpetualProps key, dicLicChannel.Item(key), arrProd, iPos, iPosOsppLic, iPosOsppLicXml
        Case "KMS_ClientC2R","KMS_Client_AE","KMS_Client","KMS_Host","KMS_Automation"
            AddSPPVlProps key, dicLicChannel.Item(key), arrProd, iPos, iPosOsppLic, iPosOsppLicXml
        End Select
    Next
End Sub 'AddLicenseData

'-------------------------------------------------------------------------------
'   AddSPPSubscriptionProps
'
'   Returns a dictionary object with Office Subscription license properties
'   for a given ACID available in the local SPP license store
'-------------------------------------------------------------------------------
Sub AddSPPSubscriptionProps (sAcid, sType, arrProd, iPos, iPosOsppLic, iPosOsppLicXml)
    On Error Resume Next
    Dim Prop, Lic
    Dim dicSppSubscriptionProps

    Set dicSppSubscriptionProps = CreateObject("Scripting.Dictionary")

    If iVersionNt > 601 Then 
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, GracePeriodRemaining, LicenseFamily, ProductKeyID, ProductKeyID2 FROM SoftwareLicensingProduct WHERE ID = '" & sAcid & "'")
    Else
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, GracePeriodRemaining, LicenseFamily, ProductKeyID, ProductKeyID2 FROM OfficeSoftwareProtectionProduct WHERE ID = '" & sAcid & "'")
    End If

    For Each Lic in Ospp
        AddSPPCommonFields dicSppSubscriptionProps, Lic
        AddToDic dicSppSubscriptionProps, "GracePeriodRemaining", CInt(Lic.GracePeriodRemaining / 1440) & " days"
        AddSPPAdditionalProps dicSppSubscriptionProps, Lic, sType
    Next
    arrProd(iPos, iPosOsppLicXml) = GetLicXmlString(arrProd(iPos, iPosOsppLicXml), dicSppSubscriptionProps)
    arrProd(iPos, iPosOsppLic) = GetLicTxtString(arrProd(iPos, iPosOsppLic), dicSppSubscriptionProps)
End Sub 'AddSPPSubscriptionProps

'-------------------------------------------------------------------------------
'   AddSPPPerpetualProps
'
'   Returns a dictionary object with Office Perpetual license properties
'   for a given ACID available in the local SPP license store
'-------------------------------------------------------------------------------
Sub AddSPPPerpetualProps (sAcid, sType, arrProd, iPos, iPosOsppLic, iPosOsppLicXml)
    On Error Resume Next
    Dim Prop, Lic
    Dim dicSppPerpetualProps

    Set dicSppPerpetualProps = CreateObject("Scripting.Dictionary")

    If iVersionNt > 601 Then 
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, LicenseFamily, ProductKeyID, ProductKeyID2 FROM SoftwareLicensingProduct WHERE ID = '" & sAcid & "'")
    Else
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, LicenseFamily, ProductKeyID, ProductKeyID2 FROM OfficeSoftwareProtectionProduct WHERE ID = '" & sAcid & "'")
    End If

    For Each Lic in Ospp
        AddSPPCommonFields dicSppPerpetualProps, Lic
        AddSPPAdditionalProps dicSppPerpetualProps, Lic, sType
    Next
    arrProd(iPos, iPosOsppLicXml) = GetLicXmlString(arrProd(iPos, iPosOsppLicXml), dicSppPerpetualProps)
    arrProd(iPos, iPosOsppLic) = GetLicTxtString(arrProd(iPos, iPosOsppLic), dicSppPerpetualProps)
End Sub 'AddSPPPerpetualProps

'-------------------------------------------------------------------------------
'   AddSPPVlProps
'
'   Returns a dictionary object with Office Perpetual license properties
'   for a given ACID available in the local SPP license store
'-------------------------------------------------------------------------------
Sub AddSPPVlProps (sAcid, sType, arrProd, iPos, iPosOsppLic, iPosOsppLicXml)
    On Error Resume Next
    Dim Prop, Lic
    Dim iVLActType
    Dim dicSppVLProps
    Dim fVLActTypeEnabled

    Set dicSppVLProps = CreateObject("Scripting.Dictionary")
    fVLActTypeEnabled = False

    If iVersionNt > 601 Then 
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, GracePeriodRemaining, LicenseFamily, ProductKeyID, KeyManagementServiceLookupDomain, VLActivationType, ADActivationObjectName, ADActivationObjectDN, ADActivationCsvlkPid, ADActivationCsvlkSkuId, VLActivationTypeEnabled, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2 FROM SoftwareLicensingProduct WHERE ID = '" & sAcid & "'")
        fVLActTypeEnabled = True
    Else
        Set OSpp = oWmiLocal.ExecQuery("SELECT ID, ApplicationId, PartialProductKey, Description, Name, LicenseStatus, LicenseStatusReason, GracePeriodRemaining, LicenseFamily, ProductKeyID, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, VLActivationInterval, VLRenewalInterval, KeyManagementServiceMachine, KeyManagementServicePort, ProductKeyID2 FROM OfficeSoftwareProtectionProduct WHERE ID = '" & sAcid & "'")
    End If

    For Each Lic in Ospp
        AddSPPCommonFields dicSppVLProps, Lic
        AddToDic dicSppVLProps, "GracePeriodRemaining", CInt(Lic.GracePeriodRemaining / 1440) & " days"
        AddSPPAdditionalProps dicSppVLProps, Lic, sType
        If fVLActTypeEnabled Then 
            iVLActType = Lic.VLActivationTypeEnabled
            If iVLActType = 0 Then
                If Lic.VLActivationType > 0 Then iVLActType = Lic.VLActivationType
            End If
            Select Case iVLActType
            Case 1 'ADBA
                AddToDic dicSppVLProps, "VLActivationTypeEnabled", iVLActType & " (ADBA)"
                AddToDic dicSppVLProps, "ADActivationObjectName", Lic.ADActivationObjectName
                AddToDic dicSppVLProps, "ADActivationObjectDN", Lic.ADActivationObjectDN
                AddToDic dicSppVLProps, "ADActivationCsvlkPid", Lic.ADActivationCsvlkPid
                AddToDic dicSppVLProps, "ADActivationCsvlkSkuId", Lic.ADActivationCsvlkSkuId
            Case 2 'KMS
                AddToDic dicSppVLProps, "VLActivationTypeEnabled", iVLActType & " (KMS)"
                AddToDic dicSppVLProps, "DiscoveredKeyManagementServiceMachineName", Lic.DiscoveredKeyManagementServiceMachineName
                AddToDic dicSppVLProps, "DiscoveredKeyManagementServiceMachinePort", Lic.DiscoveredKeyManagementServiceMachinePort
                AddToDic dicSppVLProps, "OverrideKeyManagementServiceMachineName", Lic.KeyManagementServiceMachine
                AddToDic dicSppVLProps, "OverrideKeyManagementServiceMachinePort", Lic.KeyManagementServicePort
                AddToDic dicSppVLProps, "KeyManagementServiceLookupDomain", Lic.KeyManagementServiceLookupDomain
                AddToDic dicSppVLProps, "VLActivationInterval", CInt(Lic.VLActivationInterval / 60) & " hours"
                AddToDic dicSppVLProps, "VLRenewalInterval", CInt(Lic.VLRenewalInterval / 1440) & " days"
            Case 3 'Token
                AddToDic dicSppVLProps, "VLActivationTypeEnabled", iVLActType & " (Token)"
            Case Else 'All
                AddToDic dicSppVLProps, "VLActivationTypeEnabled", iVLActType & " (ALL)"
            End Select
        Else
            AddToDic dicSppVLProps, "DiscoveredKeyManagementServiceMachineName", Lic.DiscoveredKeyManagementServiceMachineName
            AddToDic dicSppVLProps, "DiscoveredKeyManagementServiceMachinePort", Lic.DiscoveredKeyManagementServiceMachinePort
            AddToDic dicSppVLProps, "OverrideKeyManagementServiceMachineName", Lic.KeyManagementServiceMachine
            AddToDic dicSppVLProps, "OverrideKeyManagementServiceMachinePort", Lic.KeyManagementServicePort
            AddToDic dicSppVLProps, "VLActivationInterval", CInt(Lic.VLActivationInterval / 60) & " hours"
            AddToDic dicSppVLProps, "VLRenewalInterval", CInt(Lic.VLRenewalInterval / 1440) & " days"
        End If
    Next
    arrProd(iPos, iPosOsppLicXml) = GetLicXmlString(arrProd(iPos, iPosOsppLicXml), dicSppVLProps)
    arrProd(iPos, iPosOsppLic) = GetLicTxtString(arrProd(iPos, iPosOsppLic), dicSppVLProps)
End Sub 'AddSPPVlProps

'-------------------------------------------------------------------------------
'   AddSPPCommonFields
'
'   Adds common fields to spp properties dic object
'-------------------------------------------------------------------------------
Sub AddSPPCommonFields (dicSppData, Lic)
    On Error Resume Next
    AddToDic dicSppData, "ID", Lic.ID
    AddToDic dicSppData, "ApplicationId", Lic.ApplicationId
    AddToDic dicSppData, "PartialProductKey", Lic.PartialProductKey
    AddToDic dicSppData, "Description", Lic.Description
    AddToDic dicSppData, "Name", Lic.Name
    AddToDic dicSppData, "LicenseStatus", Lic.LicenseStatus
    AddToDic dicSppData, "LicenseStatusReason", Lic.LicenseStatusReason
    AddToDic dicSppData, "LicenseFamily", Lic.LicenseFamily
    AddToDic dicSppData, "ProductKeyID", Lic.ProductKeyID
    AddToDic dicSppData, "ProductKeyID2", Lic.ProductKeyID2
End Sub 'AddSPPCommonFields

'-------------------------------------------------------------------------------
'   AddSPPAdditionalProps
'
'   Adds additional detail to SPP fields to spp dic object
'-------------------------------------------------------------------------------
Sub AddSPPAdditionalProps (dicSppData, Lic, sType)
    On Error Resume Next
    Dim sTmp

    sTmp = ""
    AddToDic dicSppData, "LicenseStatusString", GetLicenseStateString(Lic.LicenseStatus)
    AddToDic dicSppData, "LicenseErrorCode","0x" & Hex(Lic.LicenseStatusReason)
    AddToDic dicSppData, "LicenseErrorCodeExplain", GetLicErrDesc(Hex(Lic.LicenseStatusReason))
    If NOT IsNull(Lic.ProductKeyID2) Then
        sTmp = Replace ((Lic.ProductKeyID2), "-", "")
        If Len(sTmp) > 18 Then sTmp = Right (sTmp, 19)
    End If
    AddToDic dicSppData, "MachineKey", Mid (sTmp, 1, 5) & "-" & Mid (sTmp, 6, 3) & "-" & Mid (sTmp, 9, 6)
    AddToDic dicSppData, "Channel", sType
End Sub 'AddSPPAdditionalProps

'-------------------------------------------------------------------------------
'   GetLicTxtString
'
'   Adds license details as txt to dic
'-------------------------------------------------------------------------------
Function GetLicTxtString (sTxt, dicSppData)
    On Error Resume Next
    sTxt = sTxt  & "#;#" & "Active License;" & GetDicItem("Name", dicSppData)
    sTxt = sTxt  & "#;#" & "Description;" & GetDicItem("Description", dicSppData)
    sTxt = sTxt  & "#;#" & "License Family;" & GetDicItem("LicenseFamily", dicSppData)
    sTxt = sTxt  & "#;#" & "Channel;" & GetDicItem("Channel", dicSppData)
    If dicSppData.Exists("GracePeriodRemaining") Then sTxt = sTxt  & "#;#" & "Remaining Grace Period;" & GetDicItem("GracePeriodRemaining", dicSppData)
    sTxt = sTxt  & "#;#" & "License Status;" & GetDicItem("LicenseStatus", dicSppData) & " - " & GetDicItem("LicenseStatusString", dicSppData) 
    sTxt = sTxt  & "#;#" & "Licensestate MoreInfo;" & GetDicItem("LicenseErrorCode", dicSppData) & " - " & GetDicItem("LicenseErrorCodeExplain", dicSppData) 
    sTxt = sTxt  & "#;#" & "Partial ProductKey;" & GetDicItem("PartialProductKey", dicSppData)
    sTxt = sTxt  & "#;#" & "SKU ID / Acid;" & GetDicItem("ID", dicSppData)
    sTxt = sTxt  & "#;#" & "ApplicationID;" & GetDicItem("ApplicationId", dicSppData)
    sTxt = sTxt  & "#;#" & "Product ID;" & GetDicItem("ProductKeyID2", dicSppData)
    sTxt = sTxt  & "#;#" & "ProductKeyID;" & GetDicItem("ProductKeyID", dicSppData)
    sTxt = sTxt  & "#;#" & "Machine Key;" & GetDicItem("MachineKey", dicSppData)
    If dicSppData.Exists("VLActivationTypeEnabled") Then sTxt = sTxt  & "#;#" & "VL ActivationType;" & GetDicItem("VLActivationTypeEnabled", dicSppData)
    If dicSppData.Exists("ADActivationObjectName") Then sTxt = sTxt  & "#;#" & "ADActivationObjectName;" & GetDicItem("ADActivationObjectName", dicSppData)
    If dicSppData.Exists("ADActivationObjectDN") Then sTxt = sTxt  & "#;#" & "ADActivationObjectDN;" & GetDicItem("ADActivationObjectDN", dicSppData)
    If dicSppData.Exists("ADActivationCsvlkPid") Then sTxt = sTxt  & "#;#" & "ADActivationCsvlkPid;" & GetDicItem("ADActivationCsvlkPid", dicSppData)
    If dicSppData.Exists("ADActivationCsvlkSkuId") Then sTxt = sTxt  & "#;#" & "ADActivationCsvlkSkuId;" & GetDicItem("ADActivationCsvlkSkuId", dicSppData)
    If dicSppData.Exists("DiscoveredKeyManagementServiceMachineName") Then sTxt = sTxt  & "#;#" & "KMS Host Discovered;" & GetDicItem("DiscoveredKeyManagementServiceMachineName", dicSppData)
    If dicSppData.Exists("OverrideKeyManagementServiceMachineName") Then sTxt = sTxt  & "#;#" & "KMS Host Override;" & GetDicItem("OverrideKeyManagementServiceMachineName", dicSppData)
    If dicSppData.Exists("DiscoveredKeyManagementServiceMachinePort") Then sTxt = sTxt  & "#;#" & "KMS Port Discovered;" & GetDicItem("DiscoveredKeyManagementServiceMachinePort", dicSppData)
    If dicSppData.Exists("OverrideKeyManagementServiceMachinePort") Then sTxt = sTxt  & "#;#" & "KMS Port Override;" & GetDicItem("OverrideKeyManagementServiceMachinePort", dicSppData)
    If dicSppData.Exists("KeyManagementServiceLookupDomain") Then sTxt = sTxt  & "#;#" & "KMS LookupDomain;" & GetDicItem("KeyManagementServiceLookupDomain", dicSppData)
    If dicSppData.Exists("VLActivationInterval") Then sTxt = sTxt  & "#;#" & "VL ActivationInterval;" & GetDicItem("VLActivationInterval", dicSppData)
    If dicSppData.Exists("VLRenewalInterval") Then sTxt = sTxt  & "#;#" & "VL RenewalInterval;" & GetDicItem("VLRenewalInterval", dicSppData)

    GetLicTxtString = sTxt
End Function 'GetLicTxtString

'-------------------------------------------------------------------------------
'   GetLicXmlString
'
'   Adds license details as xml to dic
'-------------------------------------------------------------------------------
Function GetLicXmlString (sXml, dicSppData)
    On Error Resume Next
    Dim sTmp

    If NOT sXml = "" Then sXml = sXml & vbCrLf
    sTmp = "FALSE"
    If dicSppData.Exists("PartialProductKey") Then 
        If (Len(GetDicItem("PartialProductKey", dicSppData)) = 5) Then sTmp = "TRUE"
    End If
    sXml = sXml & "<License"
    sXml = sXml & " IsActive=" & chr(34) & sTmp & chr(34)
    sXml = sXml & " Name=" & chr(34) & GetDicItem("Name", dicSppData) & chr(34)
    sXml = sXml & " Description=" & chr(34) & GetDicItem("Description", dicSppData) & chr(34)
    sXml = sXml & " Family=" & chr(34) & GetDicItem("LicenseFamily", dicSppData) & chr(34)
    sXml = sXml & " Channel=" & chr(34) & GetDicItem("Channel", dicSppData) & chr(34)
    sXml = sXml & " Status=" & chr(34) & GetDicItem("LicenseStatus", dicSppData) & chr(34)
    sXml = sXml & " StatusString=" & chr(34) & GetDicItem("LicenseStatusString", dicSppData) & chr(34)
    sXml = sXml & " StatusCode=" & chr(34) & GetDicItem("LicenseErrorCode", dicSppData) & chr(34)
    sXml = sXml & " StatusDescription=" & chr(34) & GetDicItem("LicenseErrorCodeExplain", dicSppData) & chr(34)
    sXml = sXml & " PartialProductkey=" & chr(34) & GetDicItem("PartialProductKey", dicSppData) & chr(34)
    sXml = sXml & " ApplicationID=" & chr(34) & GetDicItem("ApplicationId", dicSppData) & chr(34)
    sXml = sXml & " ProductKeyID=" & chr(34) & GetDicItem("ProductKeyID", dicSppData) & chr(34)
    sXml = sXml & " ProductID=" & chr(34) & GetDicItem("ProductKeyID2", dicSppData) & chr(34)
    sXml = sXml & " MachineKey=" & chr(34) & GetDicItem("MachineKey", dicSppData) & chr(34)
    sXml = sXml & " SkuID=" & chr(34) & GetDicItem("ID", dicSppData) & chr(34)
    sXml = sXml & " ActivationType=" & chr(34) & GetDicItem("VLActivationTypeEnabled", dicSppData) & chr(34)
    sXml = sXml & " KmsServer=" & chr(34) & GetDicItem("DiscoveredKeyManagementServiceMachineName", dicSppData) & chr(34)
    sXml = sXml & " KmsPort=" & chr(34) & GetDicItem("KeyManagementServicePort", dicSppData) & chr(34)
    sXml = sXml & " KmsServerOverride=" & chr(34) & GetDicItem("OverrideKeyManagementServiceMachineName", dicSppData) & chr(34)
    sXml = sXml & " KmsPortOverride=" & chr(34) & GetDicItem("OverrideKeyManagementServiceMachinePort", dicSppData) & chr(34)
    sXml = sXml & " ActivationInterval=" & chr(34) & GetDicItem("VLActivationInterval", dicSppData) & chr(34)
    sXml = sXml & " RenewalInterval=" & chr(34) & GetDicItem("VLRenewalInterval", dicSppData) & chr(34) 
    sXml = sXml & " KeyManagementServiceLookupDomain=" & chr(34) & GetDicItem("KeyManagementServiceLookupDomain", dicSppData) & chr(34) 
    sXml = sXml & " ActivationObjectName=" & chr(34) & GetDicItem("ADActivationObjectName", dicSppData) & chr(34)
    sXml = sXml & " ActivationObjectDN=" & chr(34) & GetDicItem("ADActivationObjectDN", dicSppData) & chr(34)
    sXml = sXml & " ActivationObjectExtendedPID=" & chr(34) & GetDicItem("ADActivationCsvlkPid", dicSppData) & chr(34)
    sXml = sXml & " ActivationObjectActivationID=" & chr(34) & GetDicItem("ADActivationCsvlkSkuId", dicSppData) & chr(34)
    sXml = sXml & " RemainingGracePeriod=" & chr(34) & GetDicItem("GracePeriodRemaining", dicSppData) & chr(34)
    sXml = sXml & " />"
    GetLicXmlString = sXml
End Function 'GetLicXmlString

'-------------------------------------------------------------------------------
'   GetDicItem
'
'   Returns the value for an item from a dictionary object
'   or an empty string if the item does not exist
'-------------------------------------------------------------------------------
Function GetDicItem (sKey, dic)
    On Error Resume Next
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    If dic.Exists(sKey) Then
        sReturn = dic.Item(sKey)
    End If
    GetDicItem = sReturn
End Function 'GetDicItem

'-------------------------------------------------------------------------------
'   GetSPPAllInstalledLics
'
'   Returns a dictionary object with 
'   all Office licenses available in the local SPP license store
'-------------------------------------------------------------------------------
Sub GetSPPAllInstalledLics (dicInstalledLics)
    On Error Resume Next
    Dim sOfficeAppId
    Dim Lic
    Dim i

    sOfficeAppId = OFFICEAPPID
    For i = 1 To 2
        If iVersionNt > 601 Then 
            Set OSpp = oWmiLocal.ExecQuery("SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationId = '" & sOfficeAppId & "'")
        Else
            Set OSpp = oWmiLocal.ExecQuery("SELECT ID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '" & sOfficeAppId & "'")
        End If

        For Each Lic in OSpp
            AddToDic dicInstalledLics, UCase(Lic.ID), ""
        Next 'Lic
        sOfficeAppId = OFFICE14APPID
    Next 'i
End Sub 'GetSPPAllInstalledLics

'-------------------------------------------------------------------------------
'   GetSPPActiveInstalledLics
'
'   Returns a dictionary object with
'   all *active* Office licenses available in the local SPP license store
'-------------------------------------------------------------------------------
Sub GetSPPActiveInstalledLics (dicInstalledActiveLics)
    On Error Resume Next
    Dim sOfficeAppId
    Dim Lic
    Dim i

    sOfficeAppId = OFFICEAPPID
    For i = 1 to 2
        If iVersionNt > 601 Then 
            Set OSpp = oWmiLocal.ExecQuery("SELECT ID FROM SoftwareLicensingProduct WHERE ApplicationId = '" & sOfficeAppId & "' "& "AND PartialProductKey <> NULL")
        Else
            Set OSpp = oWmiLocal.ExecQuery("SELECT ID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '" & sOfficeAppId & "' "& "AND PartialProductKey <> NULL")
        End If

        For Each Lic in OSpp
            AddToDic dicInstalledActiveLics, UCase(Lic.ID), ""
        Next 'Lic
        sOfficeAppId = OFFICE14APPID
    Next 'i
End Sub 'GetSPPActiveInstalledLics
'=======================================================================================================

Function GetLicenseStateString (iState)
    On Error Resume Next
    Dim sTmp

    sTmp = ""
    Select Case iState
    Case 0 : sTmp = "UNLICENSED"
    Case 1 : sTmp = "LICENSED"
    Case 2 : sTmp = "OOB GRACE"
    Case 3 : sTmp = "OOT GRACE"
    Case 4 : sTmp = "NON GENUINE GRACE"
    Case 5 : sTmp = "NOTIFICATIONS"
    Case Else : sTmp = "UNKNOWN"
    End Select

    GetLicenseStateString = sTmp
End Function 'GetLicenseStateString

'-------------------------------------------------------------------------------
'   GetPRIdAcidDic
'
'   Return a dictionary object for PRID (ProductReleaseId) to ACID mapping
'-------------------------------------------------------------------------------
Sub GetPRIdAcidDic (dicPridAcid, sArpConfigName, iVersionMajor)
    On Error Resume Next
    Dim iCnt
    Dim sSkuList, sAcidList
    Dim arrPrid, arrAcid
    
    MapConfigNameToLicense sArpConfigName, iVersionMajor, sSkuList, sAcidList
    arrPrid = Split(sSkuList, ";")
    arrAcid = Split(sAcidList, ";")

    For iCnt = 0 To UBound(arrPrid)
        AddToDic dicPridAcid, arrAcid(iCnt), arrPrid(iCnt)
    Next
End Sub 'GetPRIdAcidDic

'-------------------------------------------------------------------------------
'   MapConfigNameToLicense
'
'   Link the ARP ConfigName to the license data
'   Writes the Sku and Acid that maps to a ProductReleaseId 
'-------------------------------------------------------------------------------
Sub MapConfigNameToLicense (sArpConfigName, iVersionMajor, sSkuList, sAcidList)
    On Error Resume Next
    Select Case UCase(sArpConfigName)
        Case UCase("Access")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Access_KMS_Client;Access_MAK;Access_OEM_Perp;Access_Retail;Access_SubPrepid;Access_Trial"
                    sAcidList = "8CE7E872-188C-4B98-9D90-F8F90B7AAD02;95AB3EC8-4106-4F9D-B632-03C019D1D23F;4B17D082-D27D-467D-9E70-40B805714C0A;4D463C2C-0505-4626-8CDB-A4DA82E2D8ED;B3DADE99-DE64-4CEC-BCC7-A584D510782A;ED826596-52C4-4A81-85BD-0F343CBC6D67"
            End Select
        Case UCase("Access2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Access2019R_Grace;Access2019R_OEM_Perp;Access2019R_Retail;Access2019R_Trial"
                    sAcidList = "195CBC52-71CC-4909-A7A3-90D745D26465;71BA37DB-38DE-4F2A-B290-1D031CE3A59E;518687BD-DC55-45B9-8FA6-F918E1082E83;35CB0CD2-0B1C-47CD-B365-6E85FC2F755A"
            End Select
        Case UCase("Access2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Access2019VL_KMS_Client_AE;Access2019VL_MAK_AE"
                    sAcidList = "9E9BCEEB-E736-4F26-88DE-763F87DCC485;385B91D6-9C2C-4A2E-86B5-F44D44A48C5F"
            End Select
        Case UCase("AccessRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "AccessR_Grace;AccessR_OEM_Perp;AccessR_Retail;AccessR_Trial"
                    sAcidList = "DFC79B08-1388-44E8-8CBE-4E4E42E9F73A;DAF2E6D5-1424-47B9-8820-42B32E4BC3A2;BFA358B0-98F1-4125-842E-585FA13032E6;BB6FE552-10FB-4C0E-A4BE-6B867E075735"
                Case 15
                    sSkuList  = "AccessR_Grace;AccessR_OEM_Perp;AccessR_Retail;AccessR_Trial"
                    sAcidList = "80C94D2C-4DDE-47AE-82D2-A2ADDE81E653;70E3A52B-E7DA-4ABE-A834-B403A5776532;AB4D047B-97CF-4126-A69F-34DF08E2F254;E2E45C14-401F-4955-B05C-B6BDA44BF303"
            End Select
        Case UCase("AccessRuntime")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "AccessRuntime_ByPass"
                    sAcidList = "745FB377-0A59-4CA9-B9A9-C359557A2C4E"
            End Select
        Case UCase("AccessRuntime2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "AccessRuntime2019R_PrepidBypass"
                    sAcidList = "22E6B96C-1011-4CD5-8B35-3C8FB6366B86"
            End Select
        Case UCase("AccessRuntimeRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "AccessRuntimeR_PrepidBypass"
                    sAcidList = "9D9FAF9E-D345-4B49-AFCE-68CB0A539C7C"
                Case 15
                    sSkuList  = "AccessRuntimeR_Bypass"
                    sAcidList = "259DE5BE-492B-44B3-9D78-9645F848F7B0"
            End Select
        Case UCase("AccessVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "AccessVL_KMS_Client;AccessVL_MAK"
                    sAcidList = "67C0FC0C-DEBA-401B-BF8B-9C8AD8395804;3B2FA33F-CD5A-43A5-BD95-F49F3F546B0B"
                Case 15
                    sSkuList  = "AccessVL_KMS_Client;AccessVL_MAK"
                    sAcidList = "6EE7622C-18D8-4005-9FB7-92DB644A279B;4374022D-56B8-48C1-9BB7-D8F2FC726343"
            End Select
        Case UCase("Excel")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Excel_KMS_Client;Excel_MAK;Excel_OEM_Perp;Excel_Retail;Excel_SubPrepid;Excel_Trial"
                    sAcidList = "CEE5D470-6E3B-4FCC-8C2B-D17428568A9F;71DC86FF-F056-40D0-8FFB-9592705C9B76;EAEED721-9715-46FC-B2F8-03EEA2EF1FE2;4EAFF0D0-C6CB-4187-94F3-C7656D49A0AA;0BE50797-9053-4F15-B9B1-2F2C5A310816;8FC26056-52E4-4519-889F-CDBEDEAC7C31"
            End Select
        Case UCase("Excel2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Excel2019R_Grace;Excel2019R_OEM_Perp;Excel2019R_Retail;Excel2019R_Trial"
                    sAcidList = "381823F8-7779-4908-AAF3-5929F6BE6FDB;74605A38-8EC0-482E-AC64-872E174E505B;C201C2B7-02A1-41A8-B496-37C72910CD4A;04CAFB94-2615-419B-825A-F259E8A19D2B"
            End Select
        Case UCase("Excel2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Excel2019VL_KMS_Client_AE;Excel2019VL_MAK_AE"
                    sAcidList = "237854E9-79FC-4497-A0C1-A70969691C6B;05CB4E1D-CC81-45D5-A769-F34B09B9B391"
            End Select
        Case UCase("ExcelRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ExcelR_Grace;ExcelR_OEM_Perp;ExcelR_Retail;ExcelR_Trial"
                    sAcidList = "FC61B360-59EE-41A2-8586-E66E7B43DB89;740A90D3-A366-49FA-9845-1FA8257ED07B;424D52FF-7AD2-4BC7-8AC6-748D767B455D;A49503A7-3084-4140-B087-6B30B258F92E"
                Case 15
                    sSkuList  = "ExcelR_Grace;ExcelR_OEM_Perp;ExcelR_Retail;ExcelR_Trial"
                    sAcidList = "360E9813-EA13-4152-B020-B1D0BBF1AC17;72D74828-C6B2-420B-AC78-68BF3A0E882C;1B1D9BD5-12EA-4063-964C-16E7E87D6E08;5290A34F-2DE8-4965-B53A-2D2976CB7B35"
            End Select
        Case UCase("ExcelVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ExcelVL_KMS_Client;ExcelVL_MAK"
                    sAcidList = "C3E65D36-141F-4D2F-A303-A842EE756A29;685062A7-6024-42E7-8C5F-6BB9E63E697F"
                Case 15
                    sSkuList  = "ExcelVL_KMS_Client;ExcelVL_MAK"
                    sAcidList = "F7461D52-7C2B-43B2-8744-EA958E0BD09A;AC1AE7FD-B949-4E04-A330-849BC40638CF"
            End Select
        Case UCase("FastServer")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "FastServer_ByPass;FastServer_BypassTrial"
                    sAcidList = "DE3337D8-860B-4E99-B54C-AC38836147CB;14DE70F9-E4DE-40DD-ABE8-65AC4ADD0C72"
            End Select
        Case UCase("FastServerIA")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "FastServerIA_ByPass;FastServerIA_BypassTrial"
                    sAcidList = "503EAC89-92BC-4134-988E-0F8658B40C5E;75FFD06C-AB06-476A-8BC3-2A31701B97EA"
            End Select
        Case UCase("FastServerIB")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "FastServerIB_ByPass;FastServerIB_BypassTrial"
                    sAcidList = "5B8E93B1-E8AE-414E-8927-180D295F8F0A;67ADC8CD-7DDB-4885-8504-C1328260CEBB"
            End Select
        Case UCase("Groove")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Groove_KMS_Client;Groove_MAK;Groove_OEM_Perp;Groove_Retail;Groove_SubPrepid;Groove_Trial"
                    sAcidList = "8947D0B8-C33B-43E1-8C56-9B674C052832;FDAD0DFA-417D-4B4F-93E4-64EA8867B7FD;209408DF-98DB-4EEB-B96B-D0B9A4B13468;7004B7F0-6407-4F45-8EAC-966E5F868BDE;84155B23-BAE0-4748-967B-40E12917B0BB;DD1E0912-6816-4DC2-A8BF-AA2971DB0E25"
            End Select
        Case UCase("GrooveRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "GrooveR_Grace;GrooveR_Retail;GrooveR_Trial"
                    sAcidList = "323277B1-D81D-4329-973E-497F413BC5D0;CFAF5356-49E3-48A8-AB3C-E729AB791250;7868FC3D-ABA3-4B5D-B4EA-419C608DAB45"
            End Select
        Case UCase("GrooveServer")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "GrooveServer_ByPass;GrooveServer_BypassTrial"
                    sAcidList = "71391382-EFB5-48B3-86BD-0450650FED7D;B13B737A-E990-4819-BC00-1A12D88963C3"
            End Select
        Case UCase("GrooveServerEMS")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "GrooveServerEMS_ByPass;GrooveServerEMS_BypassTrial"
                    sAcidList = "B5300EA3-8300-42D6-A3CD-4BFFA76397BA;7C7F5B37-6BBA-49D2-B0DE-27CC76BCE67D"
            End Select
        Case UCase("GrooveVolume")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "GrooveVL_KMS_Client;GrooveVL_MAK"
                    sAcidList = "FB4875EC-0C6B-450F-B82B-AB57D8D1677F;4825AC28-CE41-45A7-9E6E-1FED74057601"
            End Select
        Case UCase("HSExcel")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HSExcel_OEM_Perp;HSExcel_Retail;HSExcel_SubPrepid;HSExcel_Trial"
                    sAcidList = "8C7E3A91-E176-4AB3-84B9-A7E0EFB3A6DD;C3AE020C-5A71-4CC5-A27A-2A97C2D46860;C7DF3516-425A-4A84-9420-0112F3094D90;F4F25E2B-C13C-4256-8E4C-5D4D82E1B862"
            End Select
        Case UCase("HSOneNote")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HSOneNote_OEM_Perp;HSOneNote_Retail;HSOneNote_SubPrepid;HSOneNote_Trial"
                    sAcidList = "9F82274C-C0EF-4212-B8D9-97A6BFBC2DC7;25FE4611-B44D-49CC-AE87-2143D299194E;D82665D5-2D8F-46BA-ABEC-FDF06206B956;B49D9ABE-7F30-40AA-9A4C-BDE08A14832D"
            End Select
        Case UCase("HSPowerPoint")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HSPowerPoint_OEM_Perp;HSPowerPoint_Retail;HSPowerPoint_SubPrepid;HSPowerPoint_Trial"
                    sAcidList = "95FF18B9-F2CF-4291-AB2E-BC17D54AA756;D652AD8D-DA5C-4358-B928-7FB1B4DE7A7C;31E631F4-EE62-4B1F-AEB6-3B30393E0045;131E900A-EFA8-412E-BA20-CB0F4BE43054"
            End Select
        Case UCase("HSWord")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HSWord_OEM_Perp;HSWord_Retail;HSWord_SubPrepid;HSWord_Trial"
                    sAcidList = "D79A3F4F-E768-4114-8D3A-7F9F45687F67;A963D7AE-7A88-41A7-94DA-8BB5635A8AF9;C735DCC2-F5E9-4077-A72F-4B6D254DDC43;533D80CB-BF68-48DB-AB3E-165B5377599E"
            End Select
        Case UCase("HomeBusiness")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HomeBusiness_BypassTrial;HomeBusiness_OEM_Perp;HomeBusiness_OEM_Perp2;HomeBusiness_OEM_Perp3;HomeBusiness_Retail;HomeBusiness_Retail2;HomeBusiness_Retail3;HomeBusiness_SubPrepid;HomeBusiness_Trial;HomeBusiness_Trial2"
                    sAcidList = "2BEB303E-66C6-4422-B2EC-5AEA48B75EE5;F63B84D0-ED9D-4B05-99E4-19D33FD7AFBD;0EAAF923-70A2-48BD-A6F1-54CC1AA95C13;00B6BBFC-4091-4182-BB81-93A9A6DEB46A;7B7D1F17-FDCB-4820-9789-9BEC6E377821;00495466-527F-442F-A681-F36FAD813F86;4EFBD4C4-5422-434C-8C25-75DA21B9381C;B21DA2D5-50F1-4C5C-BF59-07BAA35E25BA;4790B2A5-BBF2-4C26-976F-D7736E516CCE;1DFBB6C1-0C4D-44E9-A0EA-77F59146E011"
            End Select
        Case UCase("HomeBusiness2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeBusiness2019DemoR_BypassTrial180;HomeBusiness2019R_Grace;HomeBusiness2019R_OEM_Perp;HomeBusiness2019R_OEM_Perp2;HomeBusiness2019R_OEM_Perp3;HomeBusiness2019R_OEM_Perp4;HomeBusiness2019R_Retail;HomeBusiness2019R_Trial"
                    sAcidList = "64F26796-F91A-43B1-87A3-3B2E45922D71;2C18EFFF-A05D-485C-BF41-E990736DB92E;1BA73CCC-3B8A-42B7-9936-0FB1863EA6EC;EB29320D-A377-4441-BE97-83C595A7139B;B464C2C6-4ADA-4931-9B62-B449136522D0;9A93B5EB-A24C-4178-8F91-F391F0BD3923;7FE09EEF-5EED-4733-9A60-D7019DF11CAC;C08057D3-830A-43A6-8009-4E2D018EC257"
            End Select
        Case UCase("HomeBusinessDemo")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HomeBusinessDemo_BypassTrial"
                    sAcidList = "0B1ACA01-5C25-468F-809D-DA81CB49AC3A"
            End Select
        Case UCase("HomeBusinessPipcRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeBusinessPipcDemoR_BypassTrial365;HomeBusinessPipcR_Grace;HomeBusinessPipcR_OEM_Perp;HomeBusinessPipcR_PIN"
                    sAcidList = "0C6911A0-1FD3-4425-A4B3-DA478D3BF807;D95CFA27-EDDF-4116-8D28-BBDB1E07B0A5;C02FB62E-1CD5-4E18-BA25-E0480467FFAA;E260D7A1-49DC-41EE-B275-33A544C96904"
                Case 15
                    sSkuList  = "HomeBusinessPipcR_Grace;HomeBusinessPipcR_OEM_Perp;HomeBusinessPipcR_PIN"
                    sAcidList = "D95CFA27-EDDF-4116-8D28-BBDB1E07B0A5;C02FB62E-1CD5-4E18-BA25-E0480467FFAA;E260D7A1-49DC-41EE-B275-33A544C96904"
            End Select
        Case UCase("HomeBusinessRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeBusinessDemoR_BypassTrial365;HomeBusinessR_Grace;HomeBusinessR_OEM_Perp;HomeBusinessR_OEM_Perp2;HomeBusinessR_OEM_Perp3;HomeBusinessR_OEM_Perp4;HomeBusinessR_PIN;HomeBusinessR_Retail;HomeBusinessR_Retail2;HomeBusinessR_Retail3;HomeBusinessR_Trial;HomeBusinessR_Trial2"
                    sAcidList = "2AA09C36-3E86-4580-8E8D-E67DFC13A5A6;6379F5C8-BAA4-482E-A5CC-38ADE1A57348;BEA06487-C54E-436A-B535-58AEC833F05B;45C8342D-379D-4A4F-99C5-61141D3E0260;327DAB29-B51E-4A29-9204-D83A3B353D76;91AB0C97-4CBE-4889-ACA0-F0C9394046D7;55671CEE-5384-454B-9160-975BD5441C4B;86834D00-7896-4A38-8FAE-32F20B86FA2B;AC9C8FB4-387F-4492-A919-D56577AD9861;522F8458-7D49-42FF-A93B-670F6B6176AE;3A4A2CF6-5892-484D-B3D5-FE9CC5A20D78;3E0F37CC-584B-46B1-873E-C44AD57D6605"
                Case 15
                    sSkuList  = "HomeBusinessDemoR_BypassTrial180;HomeBusinessR_Grace;HomeBusinessR_OEM_Perp;HomeBusinessR_PIN;HomeBusinessR_Retail;HomeBusinessR_Subscription;HomeBusinessR_SubTest;HomeBusinessR_SubTrial;HomeBusinessR_Trial"
                    sAcidList = "8D071DB8-CDE7-4B90-8862-E2F6B54C91BF;0900883A-7F90-4A04-831D-69B5881A0C1C;1E69B3EE-DA97-421F-BED5-ABCCE247D64E;55671CEE-5384-454B-9160-975BD5441C4B;CD256150-A898-441F-AAC0-9F8F33390E45;A2B90E7A-A797-4713-AF90-F0BECF52A1DD;F5BEB18A-6861-4625-A369-9C0A2A5F512F;92847EEE-6935-4585-817D-14DCFFE6F607;BB8DF749-885C-47D8-B33A-7E5A402EF4A3"
            End Select
        Case UCase("HomeBusinessSub")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HomeBusinessSub_SubPrepid;HomeBusinessSub_Subscription;HomeBusinessSub_Subscription2"
                    sAcidList = "19316117-30A8-4773-8FD9-7F7231F4E060;2637E47C-CD16-45A1-8FF7-B7938723FD10;EFF11B33-79B0-4D87-B05F-AE5E4EC5F209"
            End Select
        Case UCase("HomeStudent")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HomeStudent_BypassTrial;HomeStudent_OEM_Perp;HomeStudent_OEM_Perp2;HomeStudent_OEM_Perp3;HomeStudent_Retail;HomeStudent_Retail2;HomeStudent_Retail3;HomeStudent_SubPrepid;HomeStudent_Trial"
                    sAcidList = "F10D4C70-F7CC-452A-B4B8-F12E3D6F4EEC;5DBE2163-3FA9-464C-B8B7-CAADDE61E4FF;0E795CCE-5BAD-40B1-8803-CE71FB89031D;DDB12F7C-CE7E-4EE5-A01C-E6AF9EDBC020;09E2D37E-474B-4121-8626-58AD9BE5776F;1CAEF4EC-ADEC-4236-A835-882F5AFD4BF0;7B0FF49B-22DA-4C74-876F-B039616D9A4E;AFCA9E83-152D-48A8-A492-6D552E40EE8A;3850C794-B06F-4633-B02F-8AC4DF0A059F"
            End Select
        Case UCase("HomeStudent2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudent2019DemoR_BypassTrial180;HomeStudent2019R_Grace;HomeStudent2019R_OEM_Perp;HomeStudent2019R_Retail;HomeStudent2019R_Trial"
                    sAcidList = "6F04E8C9-AD8A-4147-A647-DE0071BC703F;FAFAAABC-4BC0-4EFE-A9FC-B8AA32D0E101;480DF7D5-B237-45B2-96FB-E4624B12040A;4539AA2C-5C31-4D47-9139-543A868E5741;4EBC82D3-A2FA-49F3-9159-70E8DF6D45D5"
            End Select
        Case UCase("HomeStudentARM2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentARM2019PreInstallR_OEM_ARM"
                    sAcidList = "6303D14A-AFAD-431F-8434-81052A65F575"
            End Select
        Case UCase("HomeStudentARMRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentARMPreInstallR_OEM_ARM"
                    sAcidList = "090896A0-EA98-48AC-B545-BA5DA0EB0C9C"
                Case 15
                    sSkuList  = "HomeStudentARMPreInstallR_OEM_ARM"
                    sAcidList = "1FDFB4E4-F9C9-41C4-B055-C80DAF00697D"
            End Select
        Case UCase("HomeStudentDemo")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "HomeStudentDemo_BypassTrial"
                    sAcidList = "8FC4269F-A845-4D1F-9DF0-9F499C92D9CB"
            End Select
        Case UCase("HomeStudentPlusARM2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentPlusARM2019PreInstallR_OEM_ARM"
                    sAcidList = "215C841D-FFC1-4F03-BD11-5B27B6AB64CC"
            End Select
        Case UCase("HomeStudentPlusARMRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentPlusARMPreInstallR_OEM_ARM"
                    sAcidList = "6BBE2077-01A4-4269-BF15-5BF4D8EFC0B2"
                Case 15
                    sSkuList  = "HomeStudentPlusARMPreInstallR_OEM_ARM"
                    sAcidList = "EBEF9F05-5273-404A-9253-C5E252F50555"
            End Select
        Case UCase("HomeStudentRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentDemoR_BypassTrial180;HomeStudentR_Grace;HomeStudentR_OEM_Perp;HomeStudentR_PIN;HomeStudentR_Retail;HomeStudentR_Trial;HomeStudentR_Trial2"
                    sAcidList = "5F2C45C3-232D-43A1-AE90-40CD3462E9E4;1ECBF9C7-7C43-483E-B2AD-06EF5C5744C4;8E2E934E-C319-4CCE-962E-F0546FB0AA8E;253CD5CD-F190-4A8C-80AC-0144DF8990F4;C28ACDB8-D8B3-4199-BAA4-024D09E97C99;861343B3-3190-494C-A76E-A9215C7AE254;99D89E49-3F21-4D6E-BB17-3E4EE4FE8726"
                Case 15
                    sSkuList  = "HomeStudentDemoR_BypassTrial180;HomeStudentR_Grace;HomeStudentR_OEM_Perp;HomeStudentR_PIN;HomeStudentR_Retail;HomeStudentR_Subscription;HomeStudentR_SubTest;HomeStudentR_SubTrial;HomeStudentR_Trial"
                    sAcidList = "6330BD12-06E5-478A-BDF4-0CD90C3FFE65;CE6A8540-B478-4070-9ECB-2052DD288047;BE465D55-C392-4D67-BF45-4DE7829EFC6E;253CD5CD-F190-4A8C-80AC-0144DF8990F4;98685D21-78BD-4C62-BC4F-653344A63035;F2DE350D-3028-410A-BFAE-283E00B44D0E;F3F68D9F-81E5-4C7D-9910-0FD67DC2DDF2;D46BBFF5-987C-4DAC-9A8D-069BAC23AE2A;6CA43757-B2AF-4C42-9E28-F763C023EDC8"
            End Select
        Case UCase("HomeStudentVNextRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "HomeStudentVNextR_Grace;HomeStudentVNextR_Retail;HomeStudentVNextR_Trial"
                    sAcidList = "A9833863-A86B-4DEA-A034-7FDEB7A6276B;E2127526-B60C-43E0-BED1-3C9DC3D5A468;8BDA5AA2-F2DB-4124-A513-7BE79DA4E499"
            End Select
        Case UCase("InfoPath")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "InfoPath_KMS_Client;InfoPath_MAK;InfoPath_OEM_Perp;InfoPath_Retail;InfoPath_SubPrepid;InfoPath_Trial"
                    sAcidList = "CA6B6639-4AD6-40AE-A575-14DEE07F6430;85E22450-B741-430C-A172-A37962C938AF;866AC003-01A0-49FD-A6EC-F8F56ABDCFAB;EF1DA464-01C8-43A6-91AF-E4E5713744F9;0209AC7B-8A4B-450B-92F2-B583152A2613;4F0B7650-A09D-4180-976D-76D8B31EA1B4"
            End Select
        Case UCase("InfoPathRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "InfoPathR_Grace;InfoPathR_Retail;InfoPathR_Trial"
                    sAcidList = "FBE35AC1-57A8-4D02-9A23-5F97003C37D3;44984381-406E-4A35-B1C3-E54F499556E2;EAB6228E-54F9-4DDE-A218-291EA5B16C0A"
            End Select
        Case UCase("InfoPathVolume")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "InfoPathVL_KMS_Client;InfoPathVL_MAK"
                    sAcidList = "A30B8040-D68A-423F-B0B5-9CE292EA5A8F;9E016989-4007-42A6-8051-64EB97110CF2"
            End Select
        Case UCase("KMSHost2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "KMSHost2019VL_KMS_Host"
                    sAcidList = "70512334-47B4-44DB-A233-BE5EA33B914C"
            End Select
        Case UCase("KMSHostVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "KMSHostVL_KMS_Host"
                    sAcidList = "98EBFE73-2084-4C97-932C-C0CD1643BEA7"
                Case 15
                    sSkuList  = "KMSHostVL_KMS_Host"
                    sAcidList = "2E28138A-847F-42BC-9752-61B03FFF33CD"
            End Select
        Case UCase("LyncAcademicRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "LyncAcademicR_PrepidBypass"
                    sAcidList = "9103F3CE-1084-447A-827E-D6097F68C895"
            End Select
        Case UCase("LyncEntryRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "LyncEntryR_PrepidBypass"
                    sAcidList = "FF693BF4-0276-4DDB-BB42-74EF1A0C9F4D"
            End Select
        Case UCase("LyncRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "LyncR_Grace;LyncR_Retail;LyncR_Trial"
                    sAcidList = "0EA305CE-B708-4D79-8087-D636AB0F1A4D;FADA6658-BFC6-4C4E-825A-59A89822CDA8;6A88EBA4-8B91-4553-8764-61129E1E01C7"
            End Select
        Case UCase("LyncVDINone")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "LyncVDI_Bypass"
                    sAcidList = "0FBDE535-558A-4B6E-BDF7-ED6691AA7188"
            End Select
        Case UCase("LyncVolume")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "LyncVL_KMS_Client;LyncVL_MAK"
                    sAcidList = "1B9F11E3-C85C-4E1B-BB29-879AD2C909E3;E1264E10-AFAF-4439-A98B-256DF8BB156F"
            End Select
        Case UCase("MOSS")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "MOSS_ByPass;MOSS_BypassTrial"
                    sAcidList = "3FDFBCC8-B3E4-4482-91FA-122C6432805C;B2C0B444-3914-4ACB-A0B8-7CF50A8F7AA0"
            End Select
        Case UCase("MOSSFISEnt")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "MOSSFISEnt_ByPass;MOSSFISEnt_BypassTrial"
                    sAcidList = "1F230F82-E3BA-4053-9B6D-E8F061A933FC;F48AC5F3-0E33-44D8-A6A9-A629355F3F0C"
            End Select
        Case UCase("MOSSFISEntNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MOSSFISEnt_Bypass;MOSSFISEnt_BypassTrial180"
                    sAcidList = "E4083CC2-9C14-4444-9428-99F4E1828C85;45E65598-7AE9-4943-8C4D-DE6D9A508108"
                Case 15
                    sSkuList  = "MOSSFISEnt_Bypass;MOSSFISEnt_BypassTrial180"
                    sAcidList = "B695D087-638E-439A-8BE3-57DB429D8A77;2F05F7AA-6FF7-49AB-884A-14696A7BD1A4"
            End Select
        Case UCase("MOSSFISStd")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "MOSSFISStd_ByPass;MOSSFISStd_BypassTrial"
                    sAcidList = "9C2C3D66-B43B-4C22-B105-67BDD5BB6E85;8F23C830-C19B-4249-A181-B20B6DCF7DBE"
            End Select
        Case UCase("MOSSFISStdNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MOSSFISStd_Bypass;MOSSFISStd_BypassTrial180"
                    sAcidList = "0B7586A3-C196-43B2-8195-A010CCF9A612;BBFEDDD0-412D-49D5-BDA4-3F427736DE94"
                Case 15
                    sSkuList  = "MOSSFISStd_Bypass;MOSSFISStd_BypassTrial180"
                    sAcidList = "2CD676C7-3937-4AC7-B1B5-565EDA761F13;062FDB4F-3FCB-4A53-9B36-03262671B841"
            End Select
        Case UCase("MOSSNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MOSS_Bypass;MOSS_BypassTrial180"
                    sAcidList = "68CBDF33-A8E4-4E51-B27D-1E777922395E;61E77797-5801-4361-9AC4-40DD2B07F60A"
                Case 15
                    sSkuList  = "MOSS_Bypass;MOSS_BypassTrial180"
                    sAcidList = "C5D855EE-F32B-4A1C-97A8-F0A28CE02F9C;CBF97833-C73A-4BAF-9ED3-D47B3CFF51BE"
            End Select
        Case UCase("MOSSPremium")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "MOSSPremium_ByPass;MOSSPremium_BypassTrial"
                    sAcidList = "D5595F62-449B-4061-B0B2-0CBAD410BB51;88BED06D-8C6B-4E62-AB01-546D6005FE97"
            End Select
        Case UCase("MOSSPremiumNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MOSSPremium_Bypass;MOSSPremium_BypassTrial180"
                    sAcidList = "8662EA53-55DB-4A44-B72A-5E10822F9AAB;AE423093-CA28-4C04-A566-631C39C5D186"
                Case 15
                    sSkuList  = "MOSSPremium_Bypass;MOSSPremium_BypassTrial180"
                    sAcidList = "B7D84C2B-0754-49E4-B7BE-7EE321DCE0A9;298A586A-E3C1-42F0-AFE0-4BCFDC2E7CD0"
            End Select
        Case UCase("Mondo")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Mondo_ByPass;Mondo_BypassTrial;Mondo_BypassTrial2;Mondo_KMS_Client;Mondo_KMS_Client2;Mondo_MAK;Mondo_OEM_Perp;Mondo_Retail;Mondo_Retail2;Mondo_Retail3;Mondo_SubPrepid;Mondo_Subscription;Mondo_Trial;Mondo_Trial2;Mondo_Trial3;Mondo_ByPass2"
                    sAcidList = "60CD7BD4-F23C-4E58-9D81-968591E793EA;A42E403B-B141-4FA2-9781-647604D09BFB;8C893555-CB79-4BAA-94EA-780E01C3AB15;09ED9640-F020-400A-ACD8-D7D867DFD9C2;EF3D4E49-A53D-4D81-A2B1-2CA6C2556B2C;533B656A-4425-480B-8E30-1A2358898350;F6183964-39E3-46A9-A641-734041BAFF89;14F5946A-DEBC-4716-BABC-7E2C240FEC08;4671E3DF-D7B8-4502-971F-C1F86139A29A;09AE19CA-75F8-407F-BF70-8D9F24C8D222;3D9F0DCA-BCF1-4A47-8F90-EB53C6111AFD;2C971DC3-E122-4DBD-8ADF-B5777799799A;8A730535-A343-44EB-91A2-B1650DE0A872;EFEEFEDE-44B4-44FF-84D5-01B67A380024;AB324F86-606E-42EC-93E3-FC755989FE19;5B11FD8E-E03C-45DC-B35C-95145F4706EB"
            End Select
        Case UCase("MondoRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MondoR_BypassTrial180;MondoR_ConsumerSub_Bypass30;MondoR_EnterpriseSub_Bypass30;MondoR_Grace;MondoR_KMS_Automation;MondoR_O16ConsumerPerp_Bypass30;MondoR_O16EnterpriseVL_Bypass30;MondoR_O17EnterpriseVL_Bypass30;MondoR_OEM_Perp;MondoR_Retail;MondoR_Subscription;MondoR_Subscription2;MondoR_SubTest;MondoR_SubTest2;MondoR_SubTrial;MondoR_SubTrial2;MondoR_Trial;MondoR_ViewOnly_ZeroGrace"
                    sAcidList = "E7B23B5D-BDB1-495C-B7CE-CC13652ACF2F;14D002D8-B1ED-4AB4-BE6C-5F3FA7ED2F75;AAA4ABAA-369E-4D7D-93E6-4C320A017A3C;8507F630-BEE7-4833-B054-C1CA1A43990C;E914EA6E-A5FA-4439-A394-A9BB3293CA09;A7693F85-CCA9-4878-9A32-8BE2773D9331;B31C9322-8FA5-454D-AC52-100D51EDFEE5;C7136135-DC7F-46C4-9818-EA246A597480;BF3AB6CF-1B0D-4250-BAD9-4B61ECA25316;B21367DF-9545-4F02-9F24-240691DA0E58;69EC9152-153B-471A-BF35-77EC88683EAE;37256767-8FC0-4B6A-8B92-BC785BFD9D05;9E7FE6CF-58C6-4F36-9A1B-9C0BFC447ED5;CEBCF005-6D81-4C9E-8626-5D01C48B834A;26421C15-1B53-4E53-B303-F320474618E7;EE6555BB-7A24-4287-99A0-5B8B83C5D602;D229947A-7C85-4D12-944B-7536B5B05B13;762A768D-5882-4A57-A889-C387E85C10B0"
                Case 15
                    sSkuList  = "MondoPreInstallR_OEM_ARM;MondoR_BypassTrial180;MondoR_Grace;MondoR_KMS_Automation;MondoR_OEM_Perp;MondoR_Retail;MondoR_Subscription;MondoR_SubTest;MondoR_SubTrial;MondoR_Trial"
                    sAcidList = "C86F17B4-67E0-459E-B785-F09FF54600E4;1C3432FE-153B-439B-82CC-FB46D6F5983A;E1F5F599-D875-48CA-93C4-E96C473B29E7;1DC00701-03AF-4680-B2AF-007FFC758A1F;6E2F71BC-1BA0-4620-872E-285F69C3141E;3169C8DF-F659-4F95-9CC6-3115E6596E83;69EC9152-153B-471A-BF35-77EC88683EAE;9E7FE6CF-58C6-4F36-9A1B-9C0BFC447ED5;26421C15-1B53-4E53-B303-F320474618E7;2DA73A73-4C36-467F-8A2D-B49090D983B7"
            End Select
        Case UCase("MondoVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "MondoVL_KMS_Client;MondoVL_MAK"
                    sAcidList = "9CAABCCB-61B1-4B4B-8BEC-D10A3C3AC2CE;2CD0EA7E-749F-4288-A05E-567C573B2A6C"
                Case 15
                    sSkuList  = "MondoVL_KMS_Client;MondoVL_MAK"
                    sAcidList = "DC981C6B-FC8E-420F-AA43-F8F33E5C0923;F33485A0-310B-4B72-9A0E-B1D605510DBD"
            End Select
        Case UCase("O365BusinessRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "O365BusinessDemoR_BypassTrial365;O365BusinessR_Grace;O365BusinessR_Subscription;O365BusinessR_SubTest;O365BusinessR_SubTrial"
                    sAcidList = "81D507D6-E729-436D-AD4D-FED2CBE4E566;3D0631E3-1091-416D-92A5-42F84A86D868;6337137E-7C07-4197-8986-BECE6A76FC33;742178ED-6B28-42DD-B3D7-B7C0EA78741B;812837CB-040B-4829-8E22-15C4919FEBA4"
                Case 15
                    sSkuList  = "O365BusinessR_Grace;O365BusinessR_Retail;O365BusinessR_Subscription;O365BusinessR_SubTest;O365BusinessR_SubTrial"
                    sAcidList = "3D0631E3-1091-416D-92A5-42F84A86D868;BEFEE371-A2F5-4648-85DB-A2C55FDF324C;6337137E-7C07-4197-8986-BECE6A76FC33;742178ED-6B28-42DD-B3D7-B7C0EA78741B;812837CB-040B-4829-8E22-15C4919FEBA4"
            End Select
        Case UCase("O365EduCloudRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "O365EduCloudEDUR_Grace;O365EduCloudEDUR_Subscription;O365EduCloudEDUR_SubTrial"
                    sAcidList = "81E676D5-B92C-4253-AA47-36A0EA89F5ED;2F5C71B4-5B7A-4005-BB68-F9FAC26F2EA3;B2F7AD5B-0E90-4AB7-9DA5-D829A6EA3F1E"
            End Select
        Case UCase("O365HomePremRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "O365HomePremDemoR_BypassTrial365;O365HomePremR_Grace;O365HomePremR_PIN1;O365HomePremR_PIN10;O365HomePremR_PIN100;O365HomePremR_PIN101;O365HomePremR_PIN102;O365HomePremR_PIN103;O365HomePremR_PIN104;O365HomePremR_PIN105;O365HomePremR_PIN106;O365HomePremR_PIN107;O365HomePremR_PIN108;O365HomePremR_PIN109;O365HomePremR_PIN11;O365HomePremR_PIN110;O365HomePremR_PIN111;O365HomePremR_PIN112;O365HomePremR_PIN113;O365HomePremR_PIN114;O365HomePremR_PIN115;O365HomePremR_PIN116;O365HomePremR_PIN117;O365HomePremR_PIN12;O365HomePremR_PIN13;O365HomePremR_PIN14;O365HomePremR_PIN15;O365HomePremR_PIN16;O365HomePremR_PIN17;O365HomePremR_PIN18;O365HomePremR_PIN19;O365HomePremR_PIN2;O365HomePremR_PIN21;O365HomePremR_PIN22;O365HomePremR_PIN23;O365HomePremR_PIN24;O365HomePremR_PIN25;O365HomePremR_PIN26;O365HomePremR_PIN27;O365HomePremR_PIN28;O365HomePremR_PIN29;O365HomePremR_PIN3;O365HomePremR_PIN30;O365HomePremR_PIN31;O365HomePremR_PIN32;O365HomePremR_PIN33;O365HomePremR_PIN34;O365HomePremR_PIN35;O365HomePremR_PIN36;O365HomePremR_PIN37;O365HomePremR_PIN38;O365HomePremR_PIN39;O365HomePremR_PIN4;O365HomePremR_PIN40;O365HomePremR_PIN41;O365HomePremR_PIN42;O365HomePremR_PIN43;O365HomePremR_PIN44;O365HomePremR_PIN45;O365HomePremR_PIN46;O365HomePremR_PIN47;O365HomePremR_PIN48;O365HomePremR_PIN49;O365HomePremR_PIN5;O365HomePremR_PIN50;O365HomePremR_PIN51;O365HomePremR_PIN52;O365HomePremR_PIN53;O365HomePremR_PIN54;O365HomePremR_PIN55;O365HomePremR_PIN56;O365HomePremR_PIN57;O365HomePremR_PIN58;O365HomePremR_PIN59;O365HomePremR_PIN6;O365HomePremR_PIN60;O365HomePremR_PIN61;O365HomePremR_PIN62;O365HomePremR_PIN63;O365HomePremR_PIN64;O365HomePremR_PIN65;O365HomePremR_PIN66;O365HomePremR_PIN67;O365HomePremR_PIN68;O365HomePremR_PIN69;O365HomePremR_PIN7;O365HomePremR_PIN70;O365HomePremR_PIN71;O365HomePremR_PIN72;O365HomePremR_PIN73;O365HomePremR_PIN74;O365HomePremR_PIN75;O365HomePremR_PIN76;O365HomePremR_PIN77;O365HomePremR_PIN78;O365HomePremR_PIN79;O365HomePremR_PIN8;O365HomePremR_PIN80;O365HomePremR_PIN81;O365HomePremR_PIN82;O365HomePremR_PIN83;O365HomePremR_PIN84;O365HomePremR_PIN85;O365HomePremR_PIN86;O365HomePremR_PIN87;O365HomePremR_PIN88;O365HomePremR_PIN89;O365HomePremR_PIN9;O365HomePremR_PIN90;O365HomePremR_PIN91;O365HomePremR_PIN92;O365HomePremR_PIN93;O365HomePremR_PIN94;O365HomePremR_PIN95;O365HomePremR_PIN96;O365HomePremR_PIN97;O365HomePremR_PIN98;O365HomePremR_PIN99;O365HomePremR_Subscription1;O365HomePremR_Subscription2;O365HomePremR_Subscription3;O365HomePremR_Subscription4;O365HomePremR_Subscription5;O365HomePremR_SubTest1;O365HomePremR_SubTest2;O365HomePremR_SubTest3;O365HomePremR_SubTest4;O365HomePremR_SubTest5;O365HomePremR_SubTrial1;O365HomePremR_SubTrial2;O365HomePremR_SubTrial3;O365HomePremR_SubTrial4;O365HomePremR_SubTrial5"
                    sAcidList = "7AB38A1C-A803-4CA5-8EBB-F8786C4B4FAA;D7279DD0-E175-49FE-A623-8FC2FC00AFC4;0C7137CB-463A-44CA-B6CD-F29EAE4A5201;EB778317-EA9A-4FBC-8351-F00EEC82E423;52263C53-5561-4FD5-A88D-2382945E5E8B;02BB3A66-D796-4938-94E3-1D6FF2A96B5B;E6930809-6796-466F-BC45-F0937362FDC1;6516C530-02EA-49D9-AD33-A824902B1065;AA1A1140-B8A1-4377-BFF0-4DCFF82F6ED1;5770F497-3F1D-4396-A729-18677AFE6359;003D2533-6A8C-4163-9A4E-76376B21ED15;0B40C0DA-6C2C-4BB6-9C05-EF8B30F96668;59C95D45-74A6-405E-815B-99CF2692577A;ADE257E8-4824-497F-B767-C0D06CDA3E4E;9CF873F3-5112-4C73-86C5-B94808F675E0;2D40ED06-955E-4725-80FC-850CDCBC1396;B688E64D-6305-4CCF-9981-CBF6823B3F65;65A4C65F-780B-4B57-B640-2AEA70761087;7BBE17F2-24C7-416A-8821-A7ACAF91A4AA;5D8AAAF2-3EAB-48E0-822A-6EFAE35980F7;FDA59085-BDB9-4ADC-B9D7-4B5E7DFD3EAC;AB2A78FF-5C52-494A-AE2A-99C0E481CA99;9540C768-CEDF-401C-B27A-15DD833E7053;949594B8-4FBF-44A7-BD8D-911E25BA2938;2A859A5B-3B1F-43EA-B3E3-C1531CC23363;5B7C4417-66C2-46DB-A2AF-DD1D811D8DCA;7E6F537B-AF16-43E6-B07F-F457A96F58FE;AE88E21A-D981-45A0-9813-3452F5390A1E;09322079-6043-4D33-9AB1-FFC268B8248E;E0C4D41F-B115-4D51-9E8C-63AF19B6EEB8;5917FA4B-7A37-4D70-B835-2755CF165E01;12F1FDF0-D744-48B6-9F9A-B3D7744BEEB5;A46815E5-C7E6-4F4F-8994-65BD19DC831D;D059DF6E-43A4-40EE-894D-5186F85AF9A0;E47D40AD-D9D0-4AB0-9C3D-B209E77E7AE8;58DB55C7-BB77-4C4D-A170-D7ED18A4AECE;B58FB371-5E42-46CE-927B-1A27AE89194C;D84CB3EA-CB53-466A-85C8-E8ED69D3F394;286E0A48-550E-4AFE-A223-27D510DE9E3B;A4CB27DA-2524-4998-92AE-6A836124BF9F;DAA7B9D0-F65A-482B-8AD3-474A33A05BE2;21757D20-006A-4BA0-8FDB-9DE272A11CF4;44938520-350C-4919-8E2F-E840C99AF4BF;F1997D24-27E6-4350-A9C7-50E63F85406A;7F3D4AF1-E298-4C30-8118-A0533B436C43;CB9FE322-23F2-4FC5-8569-0E63D0853072;F3449F6D-D512-4BBC-AC02-8C654845F6CB;436055A7-25A6-40FF-A124-B43A216FD369;3B694869-1556-4339-98C2-96C893040E31;7BAA6D3A-9EA9-47A8-8DFF-DC33F63CFEB3;11FC26E1-ADC8-4BEE-B42A-5C57D5FEC516;396D571C-B08F-4BEE-BE0B-28C5B640BC23;B6C0DE84-37D8-41B2-9EAF-8A80A008D58A;BCEF3379-4C3F-46D6-9922-8D3E513E75FA;C299D07D-803A-423C-88E9-029DE50D79ED;8C4F02D4-30C0-441A-860D-A68BE9EFC1EE;89B643DB-DFA5-4AF4-AEFC-43D17D98B084;D7E4AF12-6EB0-40DA-BF0E-D818BBC62AFB;183C0D5A-5AAE-45EB-B0CE-88009B8B3939;2E8E87E2-ED03-4DFB-8B40-0F73785B43FD;6B102F6C-50D4-4137-BF9F-5E446348188B;BE161169-49EC-4277-9802-6C5964D96593;D495DE19-448F-4C7E-8343-E54D7F336582;16C432FF-E13E-46A5-AB7D-A1CE7519119A;E0ECE06B-9123-4314-9570-338573410DE7;1752DF1B-8314-4790-9DF8-585D0BF1D0E3;3CF36478-D8B9-4FF4-9B42-0EDC60415E6A;CCB0F52E-312E-406E-A13A-566AF3C6B91D;634C8BC9-357C-4A70-A6B1-C2A9BC1BF9E5;B6512468-269D-4910-AD2D-A6A3D769FEBA;FEDA65D8-640B-4220-B8AE-6E221A990258;F337DD12-BF19-456B-82ED-B022B3032737;CB5E4CFF-1C96-43CA-91F8-0559DBA5456C;C2B1EC65-443F-4B2A-BF99-B46B244F0179;C9D3D19B-60A6-4213-BC72-C1A1A59DC45B;943CDA82-BDAC-4298-981F-21E97303DCCE;991A2521-DBB2-436C-82EB-C1BF93AC46FD;ED7482AE-C31E-4B62-A7CB-9291A0D616D2;63F119BD-171E-4034-B98F-A85759367616;792DE587-45FD-4F49-AD84-BFD40087792C;90984839-248D-4746-8F55-C943CAB009CF;62CCF9BD-1F60-414A-B766-79E98E55283B;F600E366-4C93-4CF0-BED5-BF2398759E00;174F331D-103E-4654-B6BD-C86D5B259A1A;DEA194EA-A396-4F91-88F2-77E501C8B9D2;B64CFCED-19A0-4C81-8B8D-A72DDC7488BA;C5B6E8E0-AD3D-4712-A5C6-FB7E7CC4DA59;A9AF7310-F6E0-40E2-B44A-076B2BA7E118;1F3BF441-3BD5-4DC5-B3DA-DB90450866F4;BDA57528-BB32-43DB-8CF0-36A32372F527;F98827AD-4B5F-48F1-B81B-FEE63DF5858E;B333EAB2-86CB-4B08-A5B8-A4C64AFC7A4B;F988DF4F-10DF-444F-909F-679BB9DF70D2;4D57374B-F5EA-4513-BB95-0456E424BCEE;9AF02C45-EE74-4400-A4E7-1251D11780A0;9459DA1D-F113-416B-AB41-A81D1C29A7A0;F8FF7F1C-1D40-4B84-9535-C51F7B3A70D0;04D62A0B-845D-48ED-A872-A064FCC56B07;5CA79892-8746-4F00-BB15-F3C799EDBCF4;E96A7A9C-162D-41E1-8C3D-D4C0929B6558;934ECDD6-7438-4464-B0FD-ED374AAB9CA6;5F1FB9D5-F241-4543-994F-36B37F888585;E5606A7D-8C29-4AD7-94B5-7427BF457616;4B9B31AA-00FA-481C-A25F-A2067F3FF24C;92B51A69-397D-4796-A5E8-D1672B349469;1AE119A6-40ED-4FA9-B52A-909841BD6E3F;FA25F074-80F9-44C1-B233-4019D3E70230;77A58083-E86D-45C1-B56E-77EC925167CF;523AABDF-9353-4151-9179-F708A2B48F6B;85133BEE-AD80-4CB4-977B-F16FD2DC3A20;3DC943FD-4E44-4FD6-A05F-8B421FF1C68A;8978F58A-BDC1-418C-AB3E-8017859FCC1E;9156E89A-1370-4079-AAA5-695DC9CEEAA2;B7CD5DC5-A2E6-4E74-BA44-F16B81B931F2;87D9FF03-50F4-4247-83E7-32ED575F2B1B;064A3A09-3B2B-441D-A86B-F3D47F2D2EC6;09775609-A49B-4F28-97AE-7F7740415C66;C9693FD7-4024-4544-BF5B-45160F7D5687;537EA5B5-7D50-4876-BD38-A53A77CACA32;52287AA3-9945-44D1-9046-7A3373666821;C85FE025-B033-4F41-936B-B121521BA1C1;B58A5943-16EA-420F-A611-7B230ACD762C;28F64A3F-CF84-46DA-B1E8-2DFA7750E491;A96F8DAE-DA54-4FAD-BDC6-108DA592707A;77F47589-2212-4E3B-AD27-A900CE813837;96353198-443E-479D-9E80-9A6D72FD1A99;B4AF11BE-5F94-4D8F-9844-CE0D5E0D8680;4C1D1588-3088-4796-B3B2-8935F3A0D886;951C0FE6-40A8-4400-8003-EEC0686FFBC4;FB3540BC-A824-4C79-83DA-6F6BD3AC6CCB;86AF9220-9D6F-4575-BEF0-619A9A6AA005;AE9D158F-9450-4A0C-8C80-DD973C58357C;95820F84-6E8C-4D22-B2CE-54953E9911BC"
                Case 15
                    sSkuList  = "O365HomePremR_Grace;O365HomePremR_PIN1;O365HomePremR_PIN10;O365HomePremR_PIN100;O365HomePremR_PIN101;O365HomePremR_PIN102;O365HomePremR_PIN103;O365HomePremR_PIN104;O365HomePremR_PIN105;O365HomePremR_PIN106;O365HomePremR_PIN107;O365HomePremR_PIN108;O365HomePremR_PIN109;O365HomePremR_PIN11;O365HomePremR_PIN110;O365HomePremR_PIN111;O365HomePremR_PIN112;O365HomePremR_PIN113;O365HomePremR_PIN114;O365HomePremR_PIN115;O365HomePremR_PIN116;O365HomePremR_PIN117;O365HomePremR_PIN12;O365HomePremR_PIN13;O365HomePremR_PIN14;O365HomePremR_PIN15;O365HomePremR_PIN16;O365HomePremR_PIN17;O365HomePremR_PIN18;O365HomePremR_PIN19;O365HomePremR_PIN2;O365HomePremR_PIN21;O365HomePremR_PIN22;O365HomePremR_PIN23;O365HomePremR_PIN24;O365HomePremR_PIN25;O365HomePremR_PIN26;O365HomePremR_PIN27;O365HomePremR_PIN28;O365HomePremR_PIN29;O365HomePremR_PIN3;O365HomePremR_PIN30;O365HomePremR_PIN31;O365HomePremR_PIN32;O365HomePremR_PIN33;O365HomePremR_PIN34;O365HomePremR_PIN35;O365HomePremR_PIN36;O365HomePremR_PIN37;O365HomePremR_PIN38;O365HomePremR_PIN39;O365HomePremR_PIN4;O365HomePremR_PIN40;O365HomePremR_PIN41;O365HomePremR_PIN42;O365HomePremR_PIN43;O365HomePremR_PIN44;O365HomePremR_PIN45;O365HomePremR_PIN46;O365HomePremR_PIN47;O365HomePremR_PIN48;O365HomePremR_PIN49;O365HomePremR_PIN5;O365HomePremR_PIN50;O365HomePremR_PIN51;O365HomePremR_PIN52;O365HomePremR_PIN53;O365HomePremR_PIN54;O365HomePremR_PIN55;O365HomePremR_PIN56;O365HomePremR_PIN57;O365HomePremR_PIN58;O365HomePremR_PIN59;O365HomePremR_PIN6;O365HomePremR_PIN60;O365HomePremR_PIN61;O365HomePremR_PIN62;O365HomePremR_PIN63;O365HomePremR_PIN64;O365HomePremR_PIN65;O365HomePremR_PIN66;O365HomePremR_PIN67;O365HomePremR_PIN68;O365HomePremR_PIN69;O365HomePremR_PIN7;O365HomePremR_PIN70;O365HomePremR_PIN71;O365HomePremR_PIN72;O365HomePremR_PIN73;O365HomePremR_PIN74;O365HomePremR_PIN75;O365HomePremR_PIN76;O365HomePremR_PIN77;O365HomePremR_PIN78;O365HomePremR_PIN79;O365HomePremR_PIN8;O365HomePremR_PIN80;O365HomePremR_PIN81;O365HomePremR_PIN82;O365HomePremR_PIN83;O365HomePremR_PIN84;O365HomePremR_PIN85;O365HomePremR_PIN86;O365HomePremR_PIN87;O365HomePremR_PIN88;O365HomePremR_PIN89;O365HomePremR_PIN9;O365HomePremR_PIN90;O365HomePremR_PIN91;O365HomePremR_PIN92;O365HomePremR_PIN93;O365HomePremR_PIN94;O365HomePremR_PIN95;O365HomePremR_PIN96;O365HomePremR_PIN97;O365HomePremR_PIN98;O365HomePremR_PIN99;O365HomePremR_Subscription1;O365HomePremR_Subscription2;O365HomePremR_Subscription3;O365HomePremR_Subscription4;O365HomePremR_Subscription5;O365HomePremR_SubTest1;O365HomePremR_SubTest2;O365HomePremR_SubTest3;O365HomePremR_SubTest4;O365HomePremR_SubTest5;O365HomePremR_SubTrial1;O365HomePremR_SubTrial2;O365HomePremR_SubTrial3;O365HomePremR_SubTrial4;O365HomePremR_SubTrial5"
                    sAcidList = "D7279DD0-E175-49FE-A623-8FC2FC00AFC4;0C7137CB-463A-44CA-B6CD-F29EAE4A5201;EB778317-EA9A-4FBC-8351-F00EEC82E423;52263C53-5561-4FD5-A88D-2382945E5E8B;02BB3A66-D796-4938-94E3-1D6FF2A96B5B;E6930809-6796-466F-BC45-F0937362FDC1;6516C530-02EA-49D9-AD33-A824902B1065;AA1A1140-B8A1-4377-BFF0-4DCFF82F6ED1;5770F497-3F1D-4396-A729-18677AFE6359;003D2533-6A8C-4163-9A4E-76376B21ED15;0B40C0DA-6C2C-4BB6-9C05-EF8B30F96668;59C95D45-74A6-405E-815B-99CF2692577A;ADE257E8-4824-497F-B767-C0D06CDA3E4E;9CF873F3-5112-4C73-86C5-B94808F675E0;2D40ED06-955E-4725-80FC-850CDCBC1396;B688E64D-6305-4CCF-9981-CBF6823B3F65;65A4C65F-780B-4B57-B640-2AEA70761087;7BBE17F2-24C7-416A-8821-A7ACAF91A4AA;5D8AAAF2-3EAB-48E0-822A-6EFAE35980F7;FDA59085-BDB9-4ADC-B9D7-4B5E7DFD3EAC;AB2A78FF-5C52-494A-AE2A-99C0E481CA99;9540C768-CEDF-401C-B27A-15DD833E7053;949594B8-4FBF-44A7-BD8D-911E25BA2938;2A859A5B-3B1F-43EA-B3E3-C1531CC23363;5B7C4417-66C2-46DB-A2AF-DD1D811D8DCA;7E6F537B-AF16-43E6-B07F-F457A96F58FE;AE88E21A-D981-45A0-9813-3452F5390A1E;09322079-6043-4D33-9AB1-FFC268B8248E;E0C4D41F-B115-4D51-9E8C-63AF19B6EEB8;5917FA4B-7A37-4D70-B835-2755CF165E01;12F1FDF0-D744-48B6-9F9A-B3D7744BEEB5;A46815E5-C7E6-4F4F-8994-65BD19DC831D;D059DF6E-43A4-40EE-894D-5186F85AF9A0;E47D40AD-D9D0-4AB0-9C3D-B209E77E7AE8;58DB55C7-BB77-4C4D-A170-D7ED18A4AECE;B58FB371-5E42-46CE-927B-1A27AE89194C;D84CB3EA-CB53-466A-85C8-E8ED69D3F394;286E0A48-550E-4AFE-A223-27D510DE9E3B;A4CB27DA-2524-4998-92AE-6A836124BF9F;DAA7B9D0-F65A-482B-8AD3-474A33A05BE2;21757D20-006A-4BA0-8FDB-9DE272A11CF4;44938520-350C-4919-8E2F-E840C99AF4BF;F1997D24-27E6-4350-A9C7-50E63F85406A;7F3D4AF1-E298-4C30-8118-A0533B436C43;CB9FE322-23F2-4FC5-8569-0E63D0853072;F3449F6D-D512-4BBC-AC02-8C654845F6CB;436055A7-25A6-40FF-A124-B43A216FD369;3B694869-1556-4339-98C2-96C893040E31;7BAA6D3A-9EA9-47A8-8DFF-DC33F63CFEB3;11FC26E1-ADC8-4BEE-B42A-5C57D5FEC516;396D571C-B08F-4BEE-BE0B-28C5B640BC23;B6C0DE84-37D8-41B2-9EAF-8A80A008D58A;BCEF3379-4C3F-46D6-9922-8D3E513E75FA;C299D07D-803A-423C-88E9-029DE50D79ED;8C4F02D4-30C0-441A-860D-A68BE9EFC1EE;89B643DB-DFA5-4AF4-AEFC-43D17D98B084;D7E4AF12-6EB0-40DA-BF0E-D818BBC62AFB;183C0D5A-5AAE-45EB-B0CE-88009B8B3939;2E8E87E2-ED03-4DFB-8B40-0F73785B43FD;6B102F6C-50D4-4137-BF9F-5E446348188B;BE161169-49EC-4277-9802-6C5964D96593;D495DE19-448F-4C7E-8343-E54D7F336582;16C432FF-E13E-46A5-AB7D-A1CE7519119A;E0ECE06B-9123-4314-9570-338573410DE7;1752DF1B-8314-4790-9DF8-585D0BF1D0E3;3CF36478-D8B9-4FF4-9B42-0EDC60415E6A;CCB0F52E-312E-406E-A13A-566AF3C6B91D;634C8BC9-357C-4A70-A6B1-C2A9BC1BF9E5;B6512468-269D-4910-AD2D-A6A3D769FEBA;FEDA65D8-640B-4220-B8AE-6E221A990258;F337DD12-BF19-456B-82ED-B022B3032737;CB5E4CFF-1C96-43CA-91F8-0559DBA5456C;C2B1EC65-443F-4B2A-BF99-B46B244F0179;C9D3D19B-60A6-4213-BC72-C1A1A59DC45B;943CDA82-BDAC-4298-981F-21E97303DCCE;991A2521-DBB2-436C-82EB-C1BF93AC46FD;ED7482AE-C31E-4B62-A7CB-9291A0D616D2;63F119BD-171E-4034-B98F-A85759367616;792DE587-45FD-4F49-AD84-BFD40087792C;90984839-248D-4746-8F55-C943CAB009CF;62CCF9BD-1F60-414A-B766-79E98E55283B;F600E366-4C93-4CF0-BED5-BF2398759E00;174F331D-103E-4654-B6BD-C86D5B259A1A;DEA194EA-A396-4F91-88F2-77E501C8B9D2;B64CFCED-19A0-4C81-8B8D-A72DDC7488BA;C5B6E8E0-AD3D-4712-A5C6-FB7E7CC4DA59;A9AF7310-F6E0-40E2-B44A-076B2BA7E118;1F3BF441-3BD5-4DC5-B3DA-DB90450866F4;BDA57528-BB32-43DB-8CF0-36A32372F527;F98827AD-4B5F-48F1-B81B-FEE63DF5858E;B333EAB2-86CB-4B08-A5B8-A4C64AFC7A4B;F988DF4F-10DF-444F-909F-679BB9DF70D2;4D57374B-F5EA-4513-BB95-0456E424BCEE;9AF02C45-EE74-4400-A4E7-1251D11780A0;9459DA1D-F113-416B-AB41-A81D1C29A7A0;F8FF7F1C-1D40-4B84-9535-C51F7B3A70D0;04D62A0B-845D-48ED-A872-A064FCC56B07;5CA79892-8746-4F00-BB15-F3C799EDBCF4;E96A7A9C-162D-41E1-8C3D-D4C0929B6558;934ECDD6-7438-4464-B0FD-ED374AAB9CA6;5F1FB9D5-F241-4543-994F-36B37F888585;E5606A7D-8C29-4AD7-94B5-7427BF457616;4B9B31AA-00FA-481C-A25F-A2067F3FF24C;92B51A69-397D-4796-A5E8-D1672B349469;1AE119A6-40ED-4FA9-B52A-909841BD6E3F;FA25F074-80F9-44C1-B233-4019D3E70230;77A58083-E86D-45C1-B56E-77EC925167CF;523AABDF-9353-4151-9179-F708A2B48F6B;85133BEE-AD80-4CB4-977B-F16FD2DC3A20;3DC943FD-4E44-4FD6-A05F-8B421FF1C68A;8978F58A-BDC1-418C-AB3E-8017859FCC1E;9156E89A-1370-4079-AAA5-695DC9CEEAA2;B7CD5DC5-A2E6-4E74-BA44-F16B81B931F2;87D9FF03-50F4-4247-83E7-32ED575F2B1B;064A3A09-3B2B-441D-A86B-F3D47F2D2EC6;09775609-A49B-4F28-97AE-7F7740415C66;C9693FD7-4024-4544-BF5B-45160F7D5687;537EA5B5-7D50-4876-BD38-A53A77CACA32;52287AA3-9945-44D1-9046-7A3373666821;C85FE025-B033-4F41-936B-B121521BA1C1;B58A5943-16EA-420F-A611-7B230ACD762C;28F64A3F-CF84-46DA-B1E8-2DFA7750E491;A96F8DAE-DA54-4FAD-BDC6-108DA592707A;77F47589-2212-4E3B-AD27-A900CE813837;96353198-443E-479D-9E80-9A6D72FD1A99;B4AF11BE-5F94-4D8F-9844-CE0D5E0D8680;4C1D1588-3088-4796-B3B2-8935F3A0D886;951C0FE6-40A8-4400-8003-EEC0686FFBC4;FB3540BC-A824-4C79-83DA-6F6BD3AC6CCB;86AF9220-9D6F-4575-BEF0-619A9A6AA005;AE9D158F-9450-4A0C-8C80-DD973C58357C;95820F84-6E8C-4D22-B2CE-54953E9911BC"
            End Select
        Case UCase("O365ProPlusRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "O365ProPlusDemoR_BypassTrial365;O365ProPlusE5R_Subscription;O365ProPlusE5R_SubTrial;O365ProPlusEDUR_Subscription;O365ProPlusEDUR_SubTrial;O365ProPlusR_Grace;O365ProPlusR_Subscription1;O365ProPlusR_Subscription2;O365ProPlusR_Subscription3;O365ProPlusR_Subscription4;O365ProPlusR_Subscription5;O365ProPlusR_SubTrial1;O365ProPlusR_SubTrial2;O365ProPlusR_SubTrial3;O365ProPlusR_SubTrial4;O365ProPlusR_SubTrial5"
                    sAcidList = "35EC6E0E-2DF4-4629-9EE3-D525E806B988;7984D9ED-81F9-4D50-913D-317ECD863065;24FC428E-A37E-4996-AC66-8EA0304A152D;CBECB6F5-DA49-4029-BE25-5945AC9750B3;B27B3D00-9A95-4FCD-A0C2-118CBD5E699B;3AD61E22-E4FE-497F-BDB1-3E51BD872173;149DBCE7-A48E-44DB-8364-A53386CD4580;E3DACC06-3BC2-4E13-8E59-8E05F3232325;A8119E32-B17C-4BD3-8950-7D1853F4B412;6E5DB8A5-78E6-4953-B793-7422351AFE88;FF02E86C-FEF0-4063-B39F-74275CDDD7C3;46D2C0BD-F912-4DDC-8E67-B90EADC3F83C;26B6A7CE-B174-40AA-A114-316AA56BA9FC;B6B47040-B38E-4BE2-BF6A-DABF0C41540A;E538D623-C066-433D-A6B7-E0708B1FADF7;DFC5A8B0-E9FD-43F7-B4CA-D63F1E749711"
                Case 15
                    sSkuList  = "O365ProPlusR_Grace;O365ProPlusR_Retail;O365ProPlusR_Subscription1;O365ProPlusR_Subscription2;O365ProPlusR_Subscription3;O365ProPlusR_Subscription4;O365ProPlusR_Subscription5;O365ProPlusR_SubTrial1;O365ProPlusR_SubTrial2;O365ProPlusR_SubTrial3;O365ProPlusR_SubTrial4;O365ProPlusR_SubTrial5"
                    sAcidList = "3AD61E22-E4FE-497F-BDB1-3E51BD872173;0C4E5E7A-B436-4776-BB89-88E4B14687E2;149DBCE7-A48E-44DB-8364-A53386CD4580;E3DACC06-3BC2-4E13-8E59-8E05F3232325;A8119E32-B17C-4BD3-8950-7D1853F4B412;6E5DB8A5-78E6-4953-B793-7422351AFE88;FF02E86C-FEF0-4063-B39F-74275CDDD7C3;46D2C0BD-F912-4DDC-8E67-B90EADC3F83C;26B6A7CE-B174-40AA-A114-316AA56BA9FC;B6B47040-B38E-4BE2-BF6A-DABF0C41540A;E538D623-C066-433D-A6B7-E0708B1FADF7;DFC5A8B0-E9FD-43F7-B4CA-D63F1E749711"
            End Select
        Case UCase("O365SmallBusPremRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "O365SmallBusPremDemoR_BypassTrial365;O365SmallBusPremR_Grace;O365SmallBusPremR_PIN;O365SmallBusPremR_Subscription1;O365SmallBusPremR_Subscription2;O365SmallBusPremR_Subscription3;O365SmallBusPremR_Subscription4;O365SmallBusPremR_Subscription5;O365SmallBusPremR_SubTrial1;O365SmallBusPremR_SubTrial2;O365SmallBusPremR_SubTrial3;O365SmallBusPremR_SubTrial4;O365SmallBusPremR_SubTrial5"
                    sAcidList = "18C10CE7-C025-4615-993C-2E0C32F38EFA;8DED1DA3-5206-4540-A862-E8473B65D742;58E55634-63D9-434A-8B16-ADE5F84737B8;BACD4614-5BEF-4A5E-BAFC-DE4C788037A2;0BC1DAE4-6158-4A1C-A893-807665B934B2;65C607D5-E542-4F09-AD0B-40D6A88B2702;4CB3D290-D527-45C2-B079-26842762FDD3;881D148A-610E-4F0A-8985-3BE4C0DB2B09;6AE65B85-E04E-4368-80A7-786D5766325E;2030A84D-D6C7-4968-8CEC-CF4737ACC337;2F756D47-E1B1-4D2A-922B-4D76F35D007D;C5A706AC-9733-4061-8242-28EB639F1B29;090506FC-50F8-4C00-B8C7-91982A2A7C99"
                Case 15
                    sSkuList  = "O365SmallBusPremR_Grace;O365SmallBusPremR_PIN;O365SmallBusPremR_Retail;O365SmallBusPremR_Subscription1;O365SmallBusPremR_Subscription2;O365SmallBusPremR_Subscription3;O365SmallBusPremR_Subscription4;O365SmallBusPremR_Subscription5;O365SmallBusPremR_SubTrial1;O365SmallBusPremR_SubTrial2;O365SmallBusPremR_SubTrial3;O365SmallBusPremR_SubTrial4;O365SmallBusPremR_SubTrial5"
                    sAcidList = "8DED1DA3-5206-4540-A862-E8473B65D742;58E55634-63D9-434A-8B16-ADE5F84737B8;7A75647F-636F-4607-8E54-E1B7D1AD8930;BACD4614-5BEF-4A5E-BAFC-DE4C788037A2;0BC1DAE4-6158-4A1C-A893-807665B934B2;65C607D5-E542-4F09-AD0B-40D6A88B2702;4CB3D290-D527-45C2-B079-26842762FDD3;881D148A-610E-4F0A-8985-3BE4C0DB2B09;6AE65B85-E04E-4368-80A7-786D5766325E;2030A84D-D6C7-4968-8CEC-CF4737ACC337;2F756D47-E1B1-4D2A-922B-4D76F35D007D;C5A706AC-9733-4061-8242-28EB639F1B29;090506FC-50F8-4C00-B8C7-91982A2A7C99"
            End Select
        Case UCase("OEM")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "OEM_ByPass"
                    sAcidList = "C1CEDA8B-C578-4D5D-A4AA-23626BE4E234"
            End Select
        Case UCase("OfficeLPK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "OfficeLPK_ByPass"
                    sAcidList = "F3329A70-BB26-4DD1-AF64-68E10C1AE635"
            End Select
        Case UCase("OfficeLPK2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OfficeLPK2019_Bypass"
                    sAcidList = "6A7D5DA3-A27C-403C-87A0-BB4044A0EFBE"
            End Select
        Case UCase("OfficeLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OfficeLPK_Bypass"
                    sAcidList = "D98FA85C-4FD0-40D4-895C-6CC11DF2899F"
                Case 15
                    sSkuList  = "OfficeLPK_Bypass"
                    sAcidList = "12275A09-FEC0-45A5-8B59-446432F13CD6"
            End Select
        Case UCase("OneNote")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "OneNote_KMS_Client;OneNote_MAK;OneNote_OEM_Perp;OneNote_Retail;OneNote_SubPrepid;OneNote_Trial"
                    sAcidList = "AB586F5C-5256-4632-962F-FEFD8B49E6F4;6860B31F-6A67-48B8-84B9-E312B3485C4B;115A5CF2-D4CF-4627-91DC-839DF666D082;3F7AA693-9A7E-44FC-9309-BB3D8E604925;698FA94F-EB99-43BE-AB8C-5A085C36936C;8CC3794C-4B71-44EA-BAAE-D95CC1D17042"
            End Select
        Case UCase("OneNoteFreeRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OneNoteFreeR_Bypass"
                    sAcidList = "436366DE-5579-4F24-96DB-3893E4400030"
                Case 15
                    sSkuList  = "OneNoteFreeR_Bypass"
                    sAcidList = "3391E125-F6E4-4B1E-899C-A25E6092D40D"
            End Select
        Case UCase("OneNoteRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OneNoteR_Grace;OneNoteR_OEM_Perp;OneNoteR_Retail;OneNoteR_Trial"
                    sAcidList = "66C66DB4-55EA-427B-90A8-B850B2B87FD0;F680A57C-24C4-4FDD-80A3-778877F63D0A;83AC4DD9-1B93-40ED-AA55-EDE25BB6AF38;04388DC4-ABCA-4924-B238-AC4273F8F74B"
                Case 15
                    sSkuList  = "OneNoteR_Grace;OneNoteR_OEM_Perp;OneNoteR_Retail;OneNoteR_Trial"
                    sAcidList = "FD97BCB1-8F3C-4185-91D2-A2AB7EE278C2;BA6BA8B7-2A1D-4659-BA45-4BAC007B1698;8B524BCC-67EA-4876-A509-45E46F6347E8;B6DDD089-E96F-43F7-9A77-440E6F3CD38E"
            End Select
        Case UCase("OneNoteVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OneNoteVL_KMS_Client;OneNoteVL_MAK"
                    sAcidList = "D8CACE59-33D2-4AC7-9B1B-9B72339C51C8;23B672DA-A456-4860-A8F3-E062A501D7E8"
                Case 15
                    sSkuList  = "OneNoteVL_KMS_Client;OneNoteVL_MAK"
                    sAcidList = "EFE1F3E6-AEA2-4144-A208-32AA872B6545;B067E965-7521-455B-B9F7-C740204578A2"
            End Select
        Case UCase("Outlook")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Outlook_KMS_Client;Outlook_MAK;Outlook_OEM_Perp;Outlook_Retail;Outlook_SubPrepid;Outlook_Trial"
                    sAcidList = "ECB7C192-73AB-4DED-ACF4-2399B095D0CC;A9AEABD8-63B8-4079-A28E-F531807FD6B8;8C54246A-31FE-4274-90C7-687987735848;FBF4AC36-31C8-4340-8666-79873129CF40;8FB0D83E-2BCC-43CD-871A-6AD7A06349F4;2C8ACFCA-F0D9-4CCF-BA28-2F2C47DA8BA5"
            End Select
        Case UCase("Outlook2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Outlook2019R_Grace;Outlook2019R_OEM_Perp;Outlook2019R_Retail;Outlook2019R_Trial"
                    sAcidList = "5464FE8F-BEFA-43A3-956B-3A1FFB59A280;27D85D8B-FE2F-4515-9E37-A96056788046;20E359D5-927F-47C0-8A27-38ADBDD27124;EC9BF3B9-335A-4D5B-B9C7-7C3937CFA1F0"
            End Select
        Case UCase("Outlook2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Outlook2019VL_KMS_Client_AE;Outlook2019VL_MAK_AE"
                    sAcidList = "C8F8A301-19F5-4132-96CE-2DE9D4ADBD33;92A99ED8-2923-4CB7-A4C5-31DA6B0B8CF3"
            End Select
        Case UCase("OutlookRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OutlookR_Grace;OutlookR_OEM_Perp;OutlookR_Retail;OutlookR_Trial"
                    sAcidList = "6D47BDDA-2331-4956-A1D6-E57889F326AB;6FAC4EBC-930E-4DF5-AAAE-9FF5A36A7C99;5A670809-0983-4C2D-8AAD-D3C2C5B7D5D1;2DCD3F0B-AAFB-48D7-9F5F-01E14160AC43"
                Case 15
                    sSkuList  = "OutlookR_Grace;OutlookR_OEM_Perp;OutlookR_Retail;OutlookR_Trial"
                    sAcidList = "A3A4593B-97DC-4364-9910-70202AF2D0B5;A1B2DD4A-29DD-4B74-A6FE-BF3463C00F70;12004B48-E6C8-4FFA-AD5A-AC8D4467765A;AF3A9181-8B97-4B28-B2B1-A7AC6F8C3A05"
            End Select
        Case UCase("OutlookVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "OutlookVL_KMS_Client;OutlookVL_MAK"
                    sAcidList = "EC9D9265-9D1E-4ED0-838A-CDC20F2551A1;50059979-AC6F-4458-9E79-710BCB41721A"
                Case 15
                    sSkuList  = "OutlookVL_KMS_Client;OutlookVL_MAK"
                    sAcidList = "771C3AFA-50C5-443F-B151-FF2546D863A0;8D577C50-AE5E-47FD-A240-24986F73D503"
            End Select
        Case UCase("PTK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "PTK_ByPass"
                    sAcidList = "A4F824FC-C80E-4111-9884-DCEECE7D81A1"
            End Select
        Case UCase("PTK2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PTK2019_Bypass"
                    sAcidList = "6061FDEF-2320-46DF-B9F0-33E63D003080"
            End Select
        Case UCase("PTKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PTK_Bypass"
                    sAcidList = "AD96EEE1-FD16-4D77-8A5A-91B0C9A2369A"
                Case 15
                    sSkuList  = "PTK_Bypass"
                    sAcidList = "86E17AEA-A932-42C4-8651-95DE6CB37A08"
            End Select
        Case UCase("Personal")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Personal_OEM_Perp;Personal_Retail;Personal_Retail2;Personal_SubPrepid"
                    sAcidList = "40EC9524-41D7-41D7-95C7-E5F185F92EA3;ACB51361-C0DB-4895-9497-1831C41F31A6;45AB9E9F-C1DC-4BC7-8E37-89E0C1241776;148CE971-9CCE-4DE0-9FD6-4871BEBD5666"
            End Select
        Case UCase("Personal2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Personal2019DemoR_BypassTrial180;Personal2019R_Grace;Personal2019R_OEM_Perp;Personal2019R_Retail;Personal2019R_Trial"
                    sAcidList = "D72D694A-C20B-491A-9629-BEAFC48ECBD0;C860DD35-031F-49B8-B500-24C17A564291;1810D5D6-2572-4ED1-BD96-08D77C288DCA;2747B731-0F1F-413E-A92D-386EC1277DD8;04EC8535-E9EE-4778-A3F0-098CBF761E1E"
            End Select
        Case UCase("PersonalDemo")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "PersonalDemo_BypassTrial"
                    sAcidList = "3B02A9FF-CED3-46D6-9298-A4334829748F"
            End Select
        Case UCase("PersonalPipcRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PersonalPipcDemoR_BypassTrial365;PersonalPipcR_Grace;PersonalPipcR_OEM_Perp;PersonalPipcR_PIN"
                    sAcidList = "8BF9E61F-3905-4E9B-99C4-1C3946C1FCA7;34EBB57F-0773-4DF5-8EBD-BF55D194DC69;5AAB8561-1686-43F7-9FF5-2C861DA58D17;D545CCF1-CF46-41D0-8D55-D3E8B03CDB64"
                Case 15
                    sSkuList  = "PersonalPipcR_Grace;PersonalPipcR_OEM_Perp;PersonalPipcR_PIN"
                    sAcidList = "34EBB57F-0773-4DF5-8EBD-BF55D194DC69;5AAB8561-1686-43F7-9FF5-2C861DA58D17;D545CCF1-CF46-41D0-8D55-D3E8B03CDB64"
            End Select
        Case UCase("PersonalPrepaid")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "PersonalPrepaid_Trial"
                    sAcidList = "20F986B2-048B-4810-9421-73F31CA97657"
            End Select
        Case UCase("PersonalRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PersonalDemoR_BypassTrial180;PersonalR_Grace;PersonalR_OEM_Perp;PersonalR_Retail;PersonalR_Trial"
                    sAcidList = "338E5858-1E3A-4ACD-900E-7DBEF543BB9B;B444EF2B-340F-4E88-9E32-FCB747EE5D04;E66E9AA9-D6CE-43DA-8A34-EF62ECDAE720;A9F645A1-0D6A-4978-926A-ABCB363B72A6;FBBC4AAB-247A-43DB-8CE6-69ABE89A82E5"
                Case 15
                    sSkuList  = "PersonalDemoR_BypassTrial180;PersonalR_Grace;PersonalR_OEM_Perp;PersonalR_Retail;PersonalR_Trial"
                    sAcidList = "B845985D-10AE-449A-96AF-43CEDCE9363D;A75ADD9F-3982-464D-93FE-7D3BEC07FB46;1EFD9BEC-770F-4F5D-BC66-138ACAF5019E;17E9DF2D-ED91-4382-904B-4FED6A12CAF0;621292D6-D737-40F8-A6E8-A9D0B852AEC2"
            End Select
        Case UCase("PowerPoint")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "PowerPoint_KMS_Client;PowerPoint_MAK;PowerPoint_OEM_Perp;PowerPoint_Retail;PowerPoint_SubPrepid;PowerPoint_Trial"
                    sAcidList = "45593B1D-DFB1-4E91-BBFB-2D5D0CE2227A;38252940-718C-4AA6-81A4-135398E53851;247E7706-0D68-4F56-8D78-2B8147A11CA8;133C8359-4E93-4241-8118-30BB18737EA0;1EEA4120-6699-47E9-9E5D-2305EE108BAC;1C57AD8F-60BE-4CE0-82EC-8F55AA09751F"
            End Select
        Case UCase("PowerPoint2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PowerPoint2019R_Grace;PowerPoint2019R_OEM_Perp;PowerPoint2019R_Retail;PowerPoint2019R_Trial"
                    sAcidList = "9C31B372-0652-41B9-BDCE-34EB16400E9F;BD6CCE28-11AF-484A-97FF-86422981B868;7E63CC20-BA37-42A1-822D-D5F29F33A108;D26F7324-6471-444B-A8F5-25BBB7CB2DF4"
            End Select
        Case UCase("PowerPoint2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PowerPoint2019VL_KMS_Client_AE;PowerPoint2019VL_MAK_AE"
                    sAcidList = "3131FD61-5E4F-4308-8D6D-62BE1987C92C;13C2D7BF-F10D-42EB-9E93-ABF846785434"
            End Select
        Case UCase("PowerPointRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PowerPointR_Grace;PowerPointR_OEM_Perp;PowerPointR_Retail;PowerPointR_Trial"
                    sAcidList = "E5101C42-817D-432D-9783-8C03134A51C2;17A89B14-4244-4B97-BF6B-76D586CF91D0;F32D1284-0792-49DA-9AC6-DEB2BC9C80B6;E48D4EBB-FD30-43B2-BE87-0EB2335EC2C9"
                Case 15
                    sSkuList  = "PowerPointR_Grace;PowerPointR_OEM_Perp;PowerPointR_Retail;PowerPointR_Trial"
                    sAcidList = "F3A4939A-92A7-47CC-9D05-7C7DB72DD968;AC867F78-3A34-4C70-8D80-53B6F8C4092C;31743B82-BFBC-44B6-AA12-85D42E644D5B;52C5FBED-8BFB-4CB4-8BFF-02929CD31A98"
            End Select
        Case UCase("PowerPointVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PowerPointVL_KMS_Client;PowerPointVL_MAK"
                    sAcidList = "D70B1BBA-B893-4544-96E2-B7A318091C33;9B4060C9-A7F5-4A66-B732-FAF248B7240F"
                Case 15
                    sSkuList  = "PowerPointVL_KMS_Client;PowerPointVL_MAK"
                    sAcidList = "8C762649-97D1-4953-AD27-B7E2C25B972E;E40DCB44-1D5C-4085-8E8F-943F33C4F004"
            End Select
        Case UCase("ProPlus")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProPlus_KMS_Client;ProPlus_KMS_Host;ProPlus_MAK;ProPlus_Retail;ProPlus_Retail2;ProPlus_SubPrepid;ProPlus_Trial"
                    sAcidList = "6F327760-8C5C-417C-9B61-836A98287E0C;BFE7A195-4F8F-4F0B-A622-CF13C7D16864;FDF3ECB9-B56F-43B2-A9B8-1B48B6BAE1A7;71AF7E84-93E6-4363-9B69-699E04E74071;46C84AAD-65C7-482D-B82A-1EDC52E6989A;28FE27A7-2E11-4C05-8DD0-E1F1C08DC3AE;8C5FA740-5DCA-43F9-BE1B-D0281BCF9779"
            End Select
        Case UCase("ProPlus2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProPlus2019DemoR_BypassTrial180;ProPlus2019MSDNR_Retail;ProPlus2019R_Grace;ProPlus2019R_OEM_Perp;ProPlus2019R_OEM_Perp2;ProPlus2019R_OEM_Perp3;ProPlus2019R_OEM_Perp4;ProPlus2019R_OEM_Perp5;ProPlus2019R_OEM_Perp6;ProPlus2019R_PrepidBypass;ProPlus2019R_Retail;ProPlus2019R_Trial;ProPlus2019R_Trial2"
                    sAcidList = "CBCAF071-B911-40A1-A6C1-7BECCA481910;5F472F1E-EB0A-4170-98E2-FB9E7F6FF535;52C4D79F-6E1A-45B7-B479-36B666E0A2F8;447EFFDD-0073-4716-99F3-32EF2E155766;1640C37D-E2EB-4785-A70B-3C94A2E9B07E;6BD3E147-FF80-44FD-A3A6-B8E5E4FA9ADE;0E26561D-8B88-44CC-B87E-FAF56344E5F3;4A470B50-4D01-4FE5-910A-51BB0B7EB163;CDBD81EF-BB2C-41CC-BB3E-CA6366F3F176;C3B2612D-58EE-4D9B-AD9D-7568DECB29F7;A3072B8F-ADCC-4E75-8D62-FDEB9BDFAE57;E18C439E-424A-4E83-8F1B-5F9DD021B38D;83BD3F0A-A827-46C9-8B30-3F73BF933CFE"
            End Select
        Case UCase("ProPlus2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProPlus2019VL_KMS_Client_AE;ProPlus2019VL_MAK_AE;ProPlus2019XC2RVL_KMS_ClientC2R;ProPlus2019XC2RVL_MAKC2R"
                    sAcidList = "85DD8B5F-EAA4-4AF3-A628-CCE9E77C9A03;6755C7A7-4DFE-46F5-BCE8-427BE8E9DC62;0BC88885-718C-491D-921F-6F214349E79C;0B4E6CFA-3072-434C-B702-BD1836E6887C"
            End Select
        Case UCase("ProPlusAcad")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProPlusAcad_MAK;ProPlusAcad_Retail"
                    sAcidList = "191301D3-A579-428C-B0C7-D7988500F9E3;75BB133B-F5DD-423C-8321-3BD0B50322A5"
            End Select
        Case UCase("ProPlusMSDN")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProPlusMSDN_Retail;ProPlusMSDN_Retail2"
                    sAcidList = "42CBF3F6-4D5E-49C6-991A-0D99B8429A6D;8C5EDB5D-9AA0-47A7-9416-D61C7419A60A"
            End Select
        Case UCase("ProPlusRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProPlusDemoR_BypassTrial180;ProPlusMSDNR_Retail;ProPlusR_Grace;ProPlusR_OEM_Perp;ProPlusR_OEM_Perp2;ProPlusR_OEM_Perp3;ProPlusR_OEM_Perp4;ProPlusR_OEM_Perp5;ProPlusR_OEM_Perp6;ProPlusR_Retail;ProPlusR_Trial;ProPlusR_Trial2"
                    sAcidList = "60AFA663-984D-47A6-AC9C-00346FF5E8F0;84832881-46EF-4124-8ABC-EB493CDCF78E;70D9CEB6-6DFA-4DA4-B413-18C1C3C76E2E;6C1BED1D-0273-4045-90D2-E0836F3C380B;596BF8EC-7CAB-4A98-83AE-459DB70D24E4;26B394D7-7AD7-4AAB-8FCC-6EA678395A91;5829FD99-2B17-4BE4-9814-381145E49019;339A5901-9BDE-4F48-A88D-D048A42B54B1;AA64F755-8A7B-4519-BC32-CAB66DEB92CB;DE52BD50-9564-4ADC-8FCB-A345C17F84F9;C8CE6ADC-EDE7-4CE2-8E7B-C49F462AB8C3;E1FEF7E5-6886-458C-8E45-7C1E9DAAB00C"
                Case 15
                    sSkuList  = "ProPlusDemoR_BypassTrial180;ProPlusMSDNR_Retail;ProPlusR_Grace;ProPlusR_OEM_Perp;ProPlusR_Retail;ProPlusR_Trial"
                    sAcidList = "82E42CB5-1741-46FB-8F2F-AC8414741E8D;41499869-4103-4D3B-9DA6-D07DF41B6E39;1B686580-9FB1-4B88-BFBA-EAE7C0DA31AD;BB00F022-69DA-4BE3-968D-D9BCD41CF814;064383FA-1538-491C-859B-0ECAB169A0AB;DB56DEC3-34F2-4BC5-A7B9-ECC3CC51C12A"
            End Select
        Case UCase("ProPlusSub")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProPlusSub_SubPrepid;ProPlusSub_Subscription;ProPlusSub_Subscription2"
                    sAcidList = "E98EF0C0-71C4-42CE-8305-287D8721E26C;AE28E0AB-590F-4BE3-B7F6-438DDA6C0B1C;8D1E5912-B904-40A6-ADDD-8C7482879E87"
            End Select
        Case UCase("ProPlusVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProPlusVL_KMS_Client;ProPlusVL_MAK"
                    sAcidList = "D450596F-894D-49E0-966A-FD39ED4C4C64;C47456E3-265D-47B6-8CA0-C30ABBD0CA36"
                Case 15
                    sSkuList  = "ProPlusVL_KMS_Client;ProPlusVL_MAK"
                    sAcidList = "B322DA9C-A2E2-4058-9E4E-F59A6970BD69;2B88C4F2-EA8F-43CD-805E-4D41346E18A7"
            End Select
        Case UCase("Professional")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Professional_BypassTrial;Professional_DeltaTrial;Professional_OEM_Perp;Professional_OEM_Perp2;Professional_OEM_Perp3;Professional_Retail;Professional_Retail2;Professional_Retail3;Professional_SubPrepid;Professional_Trial;Professional_Trial2"
                    sAcidList = "4365667B-8304-463E-B542-2DF8D0A73EA9;6912CCDF-557A-497C-9903-3DE6CE9FA631;1783C7A6-840C-4B33-AF05-2B1F5CD73527;7D4627B9-9467-4AA7-AE7F-892807D78D8F;7E05FC0C-7CE4-4849-BB0B-231BDF5DCA70;8B559C37-0117-413E-921B-B853AEB6E210;50AC2361-FE88-4E5E-B0B2-13ACC96CA9AE;A7971F62-61D0-4C67-ABCC-085C10CF470F;71FB05B7-19E2-4567-AF77-8F31681D39D2;DF01848D-8F9D-4589-9198-4AC51F2547F3;42122F59-2850-485E-B0C0-1AACA1C88923"
            End Select
        Case UCase("Professional2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Professional2019DemoR_BypassTrial180;Professional2019R_Grace;Professional2019R_OEM_Perp;Professional2019R_PrepidBypass;Professional2019R_Retail;Professional2019R_Trial"
                    sAcidList = "DA892D82-124E-444B-BD01-B109155CB2AE;EB972895-440E-4863-BE98-BA075B9721D0;65C693F4-3F1C-48A6-AFE6-A6EB78A65CCB;72621D09-F833-4CC5-812B-2929A06DCF4F;1717C1E0-47D3-4899-A6D3-1022DB7415E0;EDC83D99-C10D-4ADF-8D0C-98342381820B"
            End Select
        Case UCase("ProfessionalAcad")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProfessionalAcad_OEM_Perp;ProfessionalAcad_Retail;ProfessionalAcad_Trial"
                    sAcidList = "AE3ED6AE-2654-4B82-A4BA-331265BB8972;C4109E90-6C4A-44F6-B380-EF6137122F16;23037F94-D654-4F38-962F-FF5B15348630"
            End Select
        Case UCase("ProfessionalDemo")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProfessionalDemo_BypassTrial"
                    sAcidList = "2BCDDDBE-4EBE-4728-9594-625E26137761"
            End Select
        Case UCase("ProfessionalPipcRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProfessionalPipcR_Grace;ProfessionalPipcR_OEM_Perp;ProfessionalPipcR_PIN"
                    sAcidList = "2A3D3010-CDAB-4F07-8BF2-0E3639BC6B98;4E26CAC1-E15A-4467-9069-CB47B67FE191;70A633E5-AB58-487A-B077-E410B30C5EA4"
                Case 15
                    sSkuList  = "ProfessionalPipcR_Grace;ProfessionalPipcR_OEM_Perp;ProfessionalPipcR_PIN"
                    sAcidList = "2A3D3010-CDAB-4F07-8BF2-0E3639BC6B98;4E26CAC1-E15A-4467-9069-CB47B67FE191;70A633E5-AB58-487A-B077-E410B30C5EA4"
            End Select
        Case UCase("ProfessionalRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProfessionalDemoR_BypassTrial180;ProfessionalR_Grace;ProfessionalR_OEM_Perp;ProfessionalR_PIN;ProfessionalR_Retail;ProfessionalR_Trial"
                    sAcidList = "EB264A11-A2EE-4740-8AC2-3751C8859374;7A0560C5-21ED-4518-AD41-B7F870B9FD1A;54FF28C4-56AD-4168-AB75-7207F98DE5EB;12367CA7-3B7D-44FD-8765-02469CA5004E;D64EDC00-7453-4301-8428-197343FAFB16;39A1BE8C-9E7F-4A75-81F4-21CFAC7CBECB"
                Case 15
                    sSkuList  = "ProfessionalDemoR_BypassTrial180;ProfessionalR_Grace;ProfessionalR_OEM_Perp;ProfessionalR_PIN;ProfessionalR_Retail;ProfessionalR_Trial"
                    sAcidList = "42C3EF3F-AB7C-482A-8060-B2E57B25D4BA;A9419E0F-8A3F-4E58-A143-E4B4803F85D2;3178076B-5F4E-4ACE-A160-8AAE7F002944;12367CA7-3B7D-44FD-8765-02469CA5004E;44BC70E2-FB83-4B09-9082-E5557E0C2EDE;1FCA7624-3A67-4A02-9EF3-6C363E35C8CD"
            End Select
        Case UCase("ProjectLPK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectLPK_ByPass"
                    sAcidList = "288B2C51-35B5-41CA-AA5B-524DA113B4BD"
            End Select
        Case UCase("ProjectLPK2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectLPK2019_Bypass"
                    sAcidList = "F3D55933-8E0C-4CD0-AC62-F98B3666C02D"
            End Select
        Case UCase("ProjectLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectLPK_Bypass"
                    sAcidList = "D20B83AB-A959-4FA3-9C28-82C14D74285A"
                Case 15
                    sSkuList  = "ProjectLPK_Bypass"
                    sAcidList = "781038A1-800F-4DDF-AB26-70A98774F8AD"
            End Select
        Case UCase("ProjectPro")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectPro_KMS_Client;ProjectPro_MAK;ProjectPro_OEM_Perp;ProjectPro_OEM_Perp2;ProjectPro_Retail;ProjectPro_Retail2;ProjectPro_SubPrepid;ProjectPro_Trial;ProjectPro_Trial2"
                    sAcidList = "DF133FF7-BF14-4F95-AFE3-7B48E7E331EF;1CF57A59-C532-4E56-9A7D-FFA2FE94B474;FABC9393-6174-4192-B3EE-6340E16CF90D;7B8EBE34-08FC-46C5-8BFA-161B12A43E41;725714D7-D58F-4D12-9FA8-35873C6F7215;AA188B61-D3D3-443C-9DEC-5B42393EE5CB;9A13EB9C-006F-450A-9F59-4CEC1EAB88F5;694E35B9-F965-47D7-AA19-AB2783224ADF;FD2BBCED-F8DB-45BC-B4D6-AC05A47D3691"
            End Select
        Case UCase("ProjectPro2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectPro2019DemoR_BypassTrial180;ProjectPro2019MSDNR_Retail;ProjectPro2019R_Grace;ProjectPro2019R_OEM_Perp;ProjectPro2019R_PrepidBypass;ProjectPro2019R_Retail;ProjectPro2019R_Trial"
                    sAcidList = "220842AE-9681-41F2-AF5A-4DBDE01FACC1;499E61C7-900A-4882-B003-E8E106BDCB95;7D351D13-E2F4-49EA-A63C-63A85B1D5630;BC39E263-86FB-4822-AD08-CB05F71CE855;A1BFBAE1-3092-4B33-BF77-0681BE6807FE;0D270EF7-5AAF-4370-A372-BC806B96ADB7;F6C61FE0-58F0-49E0-9D79-12F5130B48AC"
            End Select
        Case UCase("ProjectPro2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectPro2019VL_KMS_Client_AE;ProjectPro2019VL_MAK_AE;ProjectPro2019XC2RVL_KMS_ClientC2R;ProjectPro2019XC2RVL_MAKC2R"
                    sAcidList = "2CA2BF3F-949E-446A-82C7-E25A15EC78C4;D4EBADD6-401B-40D5-ADF4-A5D4ACCD72D1;FC7C4D0C-2E85-4BB9-AFD4-01ED1476B5E9;73FC2508-0FB1-475A-BD6F-2D42FD84C62A"
            End Select
        Case UCase("ProjectProMSDN")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectProMSDN_Retail"
                    sAcidList = "47A5840C-8124-4A1F-A447-50168CD6833D"
            End Select
        Case UCase("ProjectProRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectProCO365R_Subscription;ProjectProCO365R_SubTest;ProjectProCO365R_SubTrial;ProjectProDemoR_BypassTrial180;ProjectProMSDNR_Retail;ProjectProO365R_Subscription;ProjectProO365R_SubTest;ProjectProO365R_SubTrial;ProjectProR_Grace;ProjectProR_OEM_Perp;ProjectProR_Retail;ProjectProR_Retail2;ProjectProR_Trial"
                    sAcidList = "4031E9F3-E451-4E18-BEDD-91CAA3690C91;AE1310F8-2F53-4994-93A3-C61502E91D04;2B456DC8-145A-4D2D-9C72-703D1CD8C50E;13104647-EE01-4016-AFAF-BBA564215D6E;2CB19A15-BAB2-4FCB-ACEE-4BDE5BE207A5;2F72340C-B555-418D-8B46-355944FE66B8;539165C6-09E3-4F4B-9C29-EEC86FDF545F;330A4ACE-9CC1-4AF5-8D36-8D0681194618;CA5B3EEA-C055-4ACF-BC78-187DB21C7DB5;2A9E12E5-48A2-497C-98D3-BE7170D86B13;0F42F316-00B1-48C5-ADA4-2F52B5720AD0;067DABCC-B31D-4350-96A6-4FCAD79CFBB7;AEEDF8F7-8832-41B1-A9C8-13F2991A371C"
                Case 15
                    sSkuList  = "ProjectProCO365R_Subscription;ProjectProCO365R_SubTest;ProjectProCO365R_SubTrial;ProjectProDemoR_BypassTrial180;ProjectProMSDNR_Retail;ProjectProO365R_Subscription;ProjectProO365R_SubTest;ProjectProO365R_SubTrial;ProjectProR_Grace;ProjectProR_OEM_Perp;ProjectProR_Retail;ProjectProR_Trial"
                    sAcidList = "4031E9F3-E451-4E18-BEDD-91CAA3690C91;AE1310F8-2F53-4994-93A3-C61502E91D04;2B456DC8-145A-4D2D-9C72-703D1CD8C50E;40E9B240-879B-4BA4-BFB0-2E94CA998EEB;A1D1EFF9-1301-4709-BC2C-B6FCE5C158D8;2F72340C-B555-418D-8B46-355944FE66B8;539165C6-09E3-4F4B-9C29-EEC86FDF545F;330A4ACE-9CC1-4AF5-8D36-8D0681194618;AE7B1E26-3AEE-4FE3-9C5B-88F05E36CD34;3B5C7D36-1314-41C0-B405-15C558435F7D;F2435DE4-5FC0-4E5B-AC97-34F515EC5EE7;41937580-5DDD-4806-9089-5266D567219F"
            End Select
        Case UCase("ProjectProSub")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectProSub_SubPrepid;ProjectProSub_Subscription"
                    sAcidList = "4D06F72E-FD50-4BC2-A24B-D448D7F17EF2;3047CDE0-03E2-4BAE-ABC9-40AD640B418D"
            End Select
        Case UCase("ProjectProVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectProVL_KMS_Client;ProjectProVL_MAK"
                    sAcidList = "4F414197-0FC2-4C01-B68A-86CBB9AC254C;82F502B5-B0B0-4349-BD2C-C560DF85B248"
                Case 15
                    sSkuList  = "ProjectProVL_KMS_Client;ProjectProVL_MAK"
                    sAcidList = "4A5D124A-E620-44BA-B6FF-658961B33B9A;ED34DC89-1C27-4ECD-8B2F-63D0F4CEDC32"
            End Select
        Case UCase("ProjectProXVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectProXC2RVL_KMS_ClientC2R;ProjectProXC2RVL_MAKC2R"
                    sAcidList = "829B8110-0E6F-4349-BCA4-42803577788D;16728639-A9AB-4994-B6D8-F81051E69833"
            End Select
        Case UCase("ProjectServer")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectServer_ByPass;ProjectServer_BypassTrial"
                    sAcidList = "ED21638F-97FF-4A65-AD9B-6889B93065E2;84902853-59F6-4B20-BC7C-DE4F419FEFAD"
            End Select
        Case UCase("ProjectServerNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectServer_Bypass;ProjectServer_BypassTrial180"
                    sAcidList = "AEF27390-E282-44EF-B18A-649AF2928DA9;19C742D3-FAAE-4A79-89B2-4F4E5C0F5B7C"
                Case 15
                    sSkuList  = "ProjectServer_Bypass;ProjectServer_BypassTrial180"
                    sAcidList = "35466B1A-B17B-4DFB-A703-F74E2A1F5F5E;BC7BAF08-4D97-462C-8411-341052402E71"
            End Select
        Case UCase("ProjectStd")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ProjectStd_KMS_Client;ProjectStd_MAK;ProjectStd_MAK2;ProjectStd_OEM_Perp;ProjectStd_Retail;ProjectStd_SubPrepid;ProjectStd_Trial"
                    sAcidList = "5DC7BF61-5EC9-4996-9CCB-DF806A2D0EFE;11B39439-6B93-4642-9570-F2EB81BE2238;B6E9FAE1-1A0E-4C61-99D0-4AF068915378;8B1F0A02-07D6-411C-9FC7-9CAA3F86D1FE;688F6589-2BD9-424E-A152-B13F36AA6DE1;F4C9C7E4-8C96-4513-ADA3-0A514B3AC5CF;F510F8DE-4325-4461-BD33-571EDBE0A933"
            End Select
        Case UCase("ProjectStd2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectStd2019R_Grace;ProjectStd2019R_OEM_Perp;ProjectStd2019R_Retail"
                    sAcidList = "086C4557-2A7A-49B2-B4CE-3D2F93F8C8CA;DEEA27A6-E540-4B6B-9124-8A5CF73FD68B;BB7FFE5F-DAF9-4B79-B107-453E1C8427B5"
            End Select
        Case UCase("ProjectStd2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectStd2019VL_KMS_Client_AE;ProjectStd2019VL_MAK_AE"
                    sAcidList = "1777F0E3-7392-4198-97EA-8AE4DE6F6381;FDAA3C03-DC27-4A8D-8CBF-C3D843A28DDC"
            End Select
        Case UCase("ProjectStdRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectStdCO365R_Subscription;ProjectStdCO365R_SubTest;ProjectStdCO365R_SubTrial;ProjectStdO365R_Subscription;ProjectStdO365R_SubTest;ProjectStdO365R_SubTrial;ProjectStdR_Grace;ProjectStdR_OEM_Perp;ProjectStdR_Retail"
                    sAcidList = "053D3F49-B913-4B33-935E-F930DECD8709;594E9433-6492-434D-BFCE-4014F55D3909;B480F090-28FE-4A67-8885-62322037A0CD;58D95B09-6AF6-453D-A976-8EF0AE0316B1;5DF83BED-7E8E-4A28-80EF-D8B0A004CF3E;75F08E14-5CF5-4D59-9CEF-DA3194B6FD24;587BECE0-14F5-4039-AA1F-7CFCE036FCCC;6185F4C9-522C-4808-8508-946D28CEA707;E9F0B3FC-962F-4944-AD06-05C10B6BCD5E"
                Case 15
                    sSkuList  = "ProjectStdCO365R_Subscription;ProjectStdCO365R_SubTest;ProjectStdCO365R_SubTrial;ProjectStdO365R_Subscription;ProjectStdO365R_SubTest;ProjectStdO365R_SubTrial;ProjectStdR_Grace;ProjectStdR_OEM_Perp;ProjectStdR_Retail"
                    sAcidList = "053D3F49-B913-4B33-935E-F930DECD8709;594E9433-6492-434D-BFCE-4014F55D3909;B480F090-28FE-4A67-8885-62322037A0CD;58D95B09-6AF6-453D-A976-8EF0AE0316B1;5DF83BED-7E8E-4A28-80EF-D8B0A004CF3E;75F08E14-5CF5-4D59-9CEF-DA3194B6FD24;6883893B-9DD9-4943-ACEB-58327AFDC194;519FBBDE-79A6-4793-8A84-57F6541579C9;5517E6A2-739B-4822-946F-7F0F1C5934B1"
            End Select
        Case UCase("ProjectStdVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectStdVL_KMS_Client;ProjectStdVL_MAK"
                    sAcidList = "DA7DDABC-3FBE-4447-9E01-6AB7440B4CD4;82E6B314-2A62-4E51-9220-61358DD230E6"
                Case 15
                    sSkuList  = "ProjectStdVL_KMS_Client;ProjectStdVL_MAK"
                    sAcidList = "427A28D1-D17C-4ABF-B717-32C780BA6F07;2B9E4A37-6230-4B42-BEE2-E25CE86C8C7A"
            End Select
        Case UCase("ProjectStdXVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ProjectStdXC2RVL_KMS_ClientC2R;ProjectStdXC2RVL_MAKC2R"
                    sAcidList = "CBBACA45-556A-4416-AD03-BDA598EAA7C8;431058F0-C059-44C5-B9E7-ED2DD46B6789"
            End Select
        Case UCase("Publisher")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Publisher_KMS_Client;Publisher_MAK;Publisher_OEM_Perp;Publisher_Retail;Publisher_SubPrepid;Publisher_Trial"
                    sAcidList = "B50C4F75-599B-43E8-8DCD-1081A7967241;3D014759-B128-4466-9018-E80F6320D9D0;4E6F61A8-989B-463C-9948-83B894540AD4;98677603-A668-4FA4-9980-3F1F05F78F69;17B7CE1A-92A9-4B59-A7D0-E872D8A9A994;1A069855-55EC-4CCF-9C45-2AC5D500F792"
            End Select
        Case UCase("Publisher2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Publisher2019R_Grace;Publisher2019R_OEM_Perp;Publisher2019R_Retail;Publisher2019R_Trial"
                    sAcidList = "A70DB9BD-3910-427E-B565-BB8DAAAAAFFF;BDD04A4F-5804-4DC4-A1FF-8BC9FD0EC4EC;F053A7C7-F342-4AB8-9526-A1D6E5105823;771BE4FF-9148-4299-86AD-1A3EC0C10B10"
            End Select
        Case UCase("Publisher2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Publisher2019VL_KMS_Client_AE;Publisher2019VL_MAK_AE"
                    sAcidList = "9D3E4CCA-E172-46F1-A2F4-1D2107051444;40055495-BE00-444E-99CC-07446729B53E"
            End Select
        Case UCase("PublisherRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PublisherR_Grace;PublisherR_OEM_Perp;PublisherR_Retail;PublisherR_Trial"
                    sAcidList = "6CB2AFA8-FFFF-49FE-91CF-5FFBB0C9FD2B;6BB0FF35-9392-4635-B209-642C331CF1EE;6E0C1D99-C72E-4968-BCB7-AB79E03E201E;FBD9E09F-6B81-4FFB-97EF-E8919A0C9E06"
                Case 15
                    sSkuList  = "PublisherR_Grace;PublisherR_OEM_Perp;PublisherR_Retail;PublisherR_Trial"
                    sAcidList = "DB8E8683-A848-473B-B2E7-D1DE4D042095;CC802B96-22A4-4A90-8757-2ABB7B74484A;C3A0814A-70A4-471F-AF37-2313A6331111;615FEE6D-2E96-487F-98B2-51A892788A70"
            End Select
        Case UCase("PublisherVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "PublisherVL_KMS_Client;PublisherVL_MAK"
                    sAcidList = "041A06CB-C5B8-4772-809F-416D03D16654;FCC1757B-5D5F-486A-87CF-C4D6DEDB6032"
                Case 15
                    sSkuList  = "PublisherVL_KMS_Client;PublisherVL_MAK"
                    sAcidList = "00C79FF1-6850-443D-BF61-71CDE0DE305F;38EA49F6-AD1D-43F1-9888-99A35D7C9409"
            End Select
        Case UCase("SPD")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "SPD_ByPass"
                    sAcidList = "B78DF69E-0966-40B1-AE85-30A5134DEDD0"
            End Select
        Case UCase("SPDRetail")
            Select Case iVersionMajor
                Case 15
                    sSkuList  = "SPDFreeR_PrepidBypass"
                    sAcidList = "BA3E3833-6A7E-445A-89D0-7802A9A68588"
            End Select
        Case UCase("SearchServer")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "SearchServer_ByPass;SearchServer_BypassTrial"
                    sAcidList = "08460AA2-A176-442C-BDCA-26928704D80B;BC4C1C97-9013-4033-A0DD-9DC9E6D6C887"
            End Select
        Case UCase("SearchServerExpress")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "SearchServerExpress_ByPass"
                    sAcidList = "1328E89E-7EC8-4F7E-809E-7E945796E511"
            End Select
        Case UCase("ServerLPK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "ServerLPK_ByPass"
                    sAcidList = "DAF85377-642C-4222-BCD8-DF5925C70AFB"
            End Select
        Case UCase("ServerLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "ServerLPK_Bypass"
                    sAcidList = "CD3F389B-747E-4730-A9BF-87349D3F737D"
                Case 15
                    sSkuList  = "ServerLPK_Bypass"
                    sAcidList = "7930E9BD-12B8-4B32-88CE-733B4A594EDF"
            End Select
        Case UCase("SkypeServiceBypassRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeServiceBypassR_PrepidBypass"
                    sAcidList = "9103F3CE-1084-447A-827E-D6097F68C895"
            End Select
        Case UCase("SkypeforBusiness2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusiness2019R_Grace;SkypeforBusiness2019R_Retail;SkypeforBusiness2019R_Trial"
                    sAcidList = "69632864-D608-472C-AD91-9B13E195C641;B639E55C-8F3E-47FE-9761-26C6A786AD6B;90BFBEDA-4F37-4D5E-8DA9-5EF0BDBC500F"
            End Select
        Case UCase("SkypeforBusiness2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusiness2019VL_KMS_Client_AE;SkypeforBusiness2019VL_MAK_AE"
                    sAcidList = "734C6C6E-B0BA-4298-A891-671772B2BD1B;15A430D4-5E3F-4E6D-8A0A-14BF3CAEE4C7"
            End Select
        Case UCase("SkypeforBusinessEntry2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessEntry2019R_PrepidBypass"
                    sAcidList = "F88CFDEC-94CE-4463-A969-037BE92BC0E7"
            End Select
        Case UCase("SkypeforBusinessEntryRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessEntryR_PrepidBypass"
                    sAcidList = "971CD368-F2E1-49C1-AEDD-330909CE18B6"
            End Select
        Case UCase("SkypeforBusinessRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessR_Grace;SkypeforBusinessR_Retail;SkypeforBusinessR_Trial"
                    sAcidList = "17C5A0EB-24EE-4EB3-91B6-6D9F9CB1ACF6;418D2B9F-B491-4D7F-84F1-49E27CC66597;BDB32189-C636-46CD-A011-FFB834B552D9"
            End Select
        Case UCase("SkypeforBusinessVDI2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessVDI2019_Bypass"
                    sAcidList = "3F456727-422A-4134-8F62-A37380017039"
            End Select
        Case UCase("SkypeforBusinessVDINone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessVDI_Bypass"
                    sAcidList = "E7851824-6844-4D85-ABE7-7CA026FB3314"
            End Select
        Case UCase("SkypeforBusinessVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "SkypeforBusinessVL_KMS_Client;SkypeforBusinessVL_MAK"
                    sAcidList = "83E04EE1-FA8D-436D-8994-D31A862CAB77;03CA3B9A-0869-4749-8988-3CBC9D9F51BB"
            End Select
        Case UCase("SmallBusBasics")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "SmallBusBasics_KMS_Client;SmallBusBasics_MAK;SmallBusBasics_OEM_Perp;SmallBusBasics_Retail;SmallBusBasics_SubPrepid;SmallBusBasics_Trial"
                    sAcidList = "EA509E87-07A1-4A45-9EDC-EBA5A39F36AF;8090771E-D41A-4482-929E-DE87F1F47E46;88B5EC99-C9D1-47F9-B1F2-3C6C63929B7B;DBE3AEE0-5183-4FF7-8142-66050173CB01;08CEF85D-8704-417E-A749-B87C7D218CAD;4519ABCF-23DB-487B-AC28-7C9EBE801716"
            End Select
        Case UCase("SmallBusBasicsMSDN")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "SmallBusBasicsMSDN_Retail"
                    sAcidList = "AF2AFE5B-55DD-4252-AF42-E6F79CC07EBC"
            End Select
        Case UCase("Standard")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Standard_KMS_Client;Standard_MAK;Standard_Retail;Standard_SubPrepid"
                    sAcidList = "9DA2A678-FB6B-4E67-AB84-60DD6A9C819A;1F76E346-E0BE-49BC-9954-70EC53A4FCFE;D3422CFB-8D8B-4EAD-99F9-EAB0CCD990D7;BC8275B7-D67A-4390-8C5E-CC57CFC74328"
            End Select
        Case UCase("Standard2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Standard2019MSDNR_Retail;Standard2019R_Grace;Standard2019R_Retail;Standard2019R_Trial"
                    sAcidList = "1C9D26C1-4001-41BB-8B90-B83677CCC80D;42419D2A-5756-47C5-9AD3-E1548E26A4EE;FDFA34DD-A472-4B85-BEE6-CF07BF0AAA1C;4CA376C1-6DEE-4AB4-B84D-0866CDD292CF"
            End Select
        Case UCase("Standard2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Standard2019VL_KMS_Client_AE;Standard2019VL_MAK_AE"
                    sAcidList = "6912A74B-A5FB-401A-BFDB-2E3AB46F4B02;BEB5065C-1872-409E-94E2-403BCFB6A878"
            End Select
        Case UCase("StandardAcad")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "StandardAcad_MAK"
                    sAcidList = "DD457678-5C3E-48E4-BC67-A89B7A3E3B44"
            End Select
        Case UCase("StandardMSDN")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "StandardMSDN_Retail"
                    sAcidList = "B6D2565C-341D-4768-AD7D-ADDBE00BB5CE"
            End Select
        Case UCase("StandardRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "StandardMSDNR_Retail;StandardR_Grace;StandardR_Retail;StandardR_Trial"
                    sAcidList = "2395220C-BE32-4E64-81F2-6D0BDDB03EB1;C7D96193-844A-44BD-93D0-76E4B785248E;4A31C291-3A12-4C64-B8AB-CD79212BE45E;6CE96173-4198-49FC-9924-CC8EEC1A7285"
                Case 15
                    sSkuList  = "StandardMSDNR_Retail;StandardR_Grace;StandardR_Retail;StandardR_Trial"
                    sAcidList = "A884AB66-DE2E-4505-BA8B-3FB9DAF149ED;370155F7-D1EC-4B8B-9FBA-EE8ACC6E0BB7;32255C0A-16B4-4CE2-B388-8A4267E219EB;603EDDF5-E827-4233-AFDE-4084B3F3B30C"
            End Select
        Case UCase("StandardVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "StandardVL_KMS_Client;StandardVL_MAK"
                    sAcidList = "DEDFA23D-6ED1-45A6-85DC-63CAE0546DE6;0ED94AAC-2234-4309-BA29-74BDBB887083"
                Case 15
                    sSkuList  = "StandardVL_KMS_Client;StandardVL_MAK"
                    sAcidList = "B13AFB38-CD79-4AE5-9F7F-EED058D750CA;A24CCA51-3D54-4C41-8A76-4031F5338CB2"
            End Select
        Case UCase("Starter")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Starter_ByPass"
                    sAcidList = "2745E581-565A-4670-AE90-6BF7C57FFE43"
            End Select
        Case UCase("StarterPrem")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "StarterPrem_SubPrepid"
                    sAcidList = "59EC6B79-F6F5-4ADD-A5A0-B755BFB77422"
            End Select
        Case UCase("Sub4")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Sub4_Subscription"
                    sAcidList = "EFF34BA1-2794-4908-9501-5190C8F2025E"
            End Select
        Case UCase("VisioLPK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "VisioLPK_ByPass"
                    sAcidList = "4A8825CE-A527-4A42-B59C-17B22CFE644F"
            End Select
        Case UCase("VisioLPK2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioLPK2019_Bypass"
                    sAcidList = "20AB14AF-1B6A-4C6B-A469-43EC0FE5AF78"
            End Select
        Case UCase("VisioLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioLPK_Bypass"
                    sAcidList = "D4639CC7-7B3A-4AC6-99C4-B0AF9ACE7B27"
                Case 15
                    sSkuList  = "VisioLPK_Bypass"
                    sAcidList = "E0132983-CA20-4390-8BBA-8FDA37E7C86B"
            End Select
        Case UCase("VisioPrem")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "VisioPrem_KMS_Client;VisioPrem_MAK;VisioPrem_OEM_Perp;VisioPrem_Retail;VisioPrem_SubPrepid;VisioPrem_Trial"
                    sAcidList = "92236105-BB67-494F-94C7-7F7A607929BD;36756CB8-8E69-4D11-9522-68899507CD6A;BB42DD2B-070C-4F5B-947A-55F56A16D4F3;66CAD568-C2DC-459D-93EC-2F3CB967EE34;3513C04B-9085-43A9-8F9A-639993C19E80;7616C283-5C5B-4054-B52C-902F03E4DCDF"
            End Select
        Case UCase("VisioPro")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "VisioPro_KMS_Client;VisioPro_MAK;VisioPro_OEM_Perp;VisioPro_OEM_Perp2;VisioPro_Retail;VisioPro_Retail2;VisioPro_SubPrepid;VisioPro_Trial"
                    sAcidList = "E558389C-83C3-4B29-ADFE-5E4D7F46C358;5980CF2B-E460-48AF-921E-0C2A79025D23;1359DCE0-0DC8-4171-8C43-BA8B9F2E5D1D;0B172E55-95AE-4C78-8C58-81AA98AB7F94;D0A97E12-08A1-4A45-ADD5-1155B204E766;0993043D-664F-4B2E-A7F1-FD92091FA81F;0EC894E8-A5A9-48DE-9463-061C4801EE8F;673EA9EA-9BC0-463F-93E5-F77655E46630"
            End Select
        Case UCase("VisioPro2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioPro2019DemoR_BypassTrial180;VisioPro2019MSDNR_Retail;VisioPro2019R_Grace;VisioPro2019R_OEM_Perp;VisioPro2019R_PrepidBypass;VisioPro2019R_Retail;VisioPro2019R_Trial"
                    sAcidList = "C726993A-CD6C-499F-B672-F36AECAFFB20;0758C976-EA87-4128-8C1C-F7AF1BAB76B9;DB2760F6-F36B-403C-A0E7-0B59CDE7427B;4EF70E76-5115-4BD2-9B12-3B0CB671AD37;2EA0D3C6-B498-4503-97A1-19AC88AE44E5;A6F69D68-5590-4E02-80B9-E7233DFF204E;D51BB17C-7E2B-45F4-A9F3-D2D444CC1E13"
            End Select
        Case UCase("VisioPro2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioPro2019VL_KMS_Client_AE;VisioPro2019VL_MAK_AE;VisioPro2019XC2RVL_KMS_ClientC2R;VisioPro2019XC2RVL_MAKC2R"
                    sAcidList = "5B5CF08F-B81A-431D-B080-3450D8620565;F41ABF81-F409-4B0D-889D-92B3E3D7D005;500F6619-EF93-4B75-BCB4-82819998A3CA;7DF5867B-E38F-459F-A590-93C1DD2EF0C6"
            End Select
        Case UCase("VisioProMSDN")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "VisioProMSDN_Retail"
                    sAcidList = "15A9D881-3184-45E0-B407-466A68A691B1"
            End Select
        Case UCase("VisioProRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioProCO365R_Subscription;VisioProCO365R_SubTest;VisioProCO365R_SubTrial;VisioProDemoR_BypassTrial180;VisioProMSDNR_Retail;VisioProO365R_Subscription;VisioProO365R_SubTest;VisioProO365R_SubTrial;VisioProR_Grace;VisioProR_OEM_Perp;VisioProR_Retail;VisioProR_Retail2;VisioProR_Trial"
                    sAcidList = "4B2E77A0-8321-42A0-B36D-66A2CDB27AF4;6C44E4BD-C6DB-474A-877E-1BB899ACA206;79EEBE53-530A-4869-B56B-B4FFE0A1B831;D0C2BC53-DB5F-4A9E-84AC-17C8862C27DA;3846A94F-3005-401E-B663-AF1CF99813D5;A56A3B37-3A35-4BBB-A036-EEE5F1898EEE;27B162B5-F5D2-40E9-8AF3-D42FF4572BD4;402DFB6D-631E-4A3F-9A5A-BE753BFD84C5;5821EC16-77A9-4404-99C8-2756DC6D4C3C;ED967494-A18E-41A9-ABB8-9030C51BA685;2DFE2075-2D04-4E43-816A-EB60BBB77574;1F7413CF-88D2-4F49-A283-E4D757D179DF;A17F9ED0-C3D4-4873-B3B8-D7E049B459EC"
                Case 15
                    sSkuList  = "VisioProCO365R_Subscription;VisioProCO365R_SubTest;VisioProCO365R_SubTrial;VisioProDemoR_BypassTrial180;VisioProMSDNR_Retail;VisioProO365R_Subscription;VisioProO365R_SubTest;VisioProO365R_SubTrial;VisioProR_Grace;VisioProR_OEM_Perp;VisioProR_Retail;VisioProR_Trial"
                    sAcidList = "4B2E77A0-8321-42A0-B36D-66A2CDB27AF4;6C44E4BD-C6DB-474A-877E-1BB899ACA206;79EEBE53-530A-4869-B56B-B4FFE0A1B831;6A9BD082-8C07-462C-8D19-18C2CAE0FB02;A491EFAD-26CE-4ED7-B722-4569BA761D14;A56A3B37-3A35-4BBB-A036-EEE5F1898EEE;27B162B5-F5D2-40E9-8AF3-D42FF4572BD4;402DFB6D-631E-4A3F-9A5A-BE753BFD84C5;024EA285-2685-48BC-87EF-79B48CC8C027;141F6878-2591-4231-A476-027432E28B2F;15D12AD4-622D-4257-976C-5EB3282FB93D;F35E39C1-A41F-47C9-A204-2CA3C4B13548"
            End Select
        Case UCase("VisioProVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioProVL_KMS_Client;VisioProVL_MAK"
                    sAcidList = "6BF301C1-B94A-43E9-BA31-D494598C47FB;295B2C03-4B1C-4221-B292-1411F468BD02"
                Case 15
                    sSkuList  = "VisioProVL_KMS_Client;VisioProVL_MAK"
                    sAcidList = "E13AC10E-75D0-4AFF-A0CD-764982CF541C;3E4294DD-A765-49BC-8DBD-CF8B62A4BD3D"
            End Select
        Case UCase("VisioProXVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioProXC2RVL_KMS_ClientC2R;VisioProXC2RVL_MAKC2R"
                    sAcidList = "B234ABE3-0857-4F9C-B05A-4DC314F85557;0594DC12-8444-4912-936A-747CA742DBDB"
            End Select
        Case UCase("VisioStd")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "VisioStd_KMS_Client;VisioStd_MAK;VisioStd_OEM_Perp;VisioStd_Retail;VisioStd_SubPrepid;VisioStd_Trial"
                    sAcidList = "9ED833FF-4F92-4F36-B370-8683A4F13275;CAB3A4C4-F31A-4C12-AFA9-A0EECC86BD95;40BECF98-1D17-43EF-989F-1D92396A2741;BA24D057-8B5F-462E-87FE-485038C68954;4CC91C85-44A8-4834-B15D-FFEA4616E4E4;A27DF0C4-AE71-4DDD-BBEB-6D6222FE3A17"
            End Select
        Case UCase("VisioStd2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioStd2019R_Grace;VisioStd2019R_OEM_Perp;VisioStd2019R_Retail"
                    sAcidList = "8497DB0B-415B-4900-A20E-ECEEC4A6D350;C613537D-13A0-4D13-939C-BC876A2F5E8C;4A582021-18C2-489F-9B3D-5186DE48F1CD"
            End Select
        Case UCase("VisioStd2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioStd2019VL_KMS_Client_AE;VisioStd2019VL_MAK_AE"
                    sAcidList = "E06D7DF3-AAD0-419D-8DFB-0AC37E2BDF39;933ED0E3-747D-48B0-9C2C-7CEB4C7E473D"
            End Select
        Case UCase("VisioStdRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioStdCO365R_Subscription;VisioStdCO365R_SubTest;VisioStdCO365R_SubTrial;VisioStdO365R_Subscription;VisioStdO365R_SubTest;VisioStdO365R_SubTrial;VisioStdR_Grace;VisioStdR_OEM_Perp;VisioStdR_Retail"
                    sAcidList = "C5A7B998-FAE7-4F08-9802-9A2216F1EC2B;1B170697-99F8-4B3C-861D-FBDBE5303A8A;2F632ABD-D62A-46C8-BF95-70186096F1ED;980F9E3E-F5A8-41C8-8596-61404ADDF677;CEED49FA-52AD-4C90-B7D4-926ECDFD7F52;2E5F5521-6EFA-4E29-804D-993C1D1D7406;AD310F32-BD8D-40AD-A882-BC7FCBAD5259;15265EDE-4272-4B92-9D47-EF584858D0AE;C76DBCBC-D71B-4F45-B5B3-B7494CB4E23E"
                Case 15
                    sSkuList  = "VisioStdCO365R_Subscription;VisioStdCO365R_SubTest;VisioStdCO365R_SubTrial;VisioStdO365R_Subscription;VisioStdO365R_SubTest;VisioStdO365R_SubTrial;VisioStdR_Grace;VisioStdR_OEM_Perp;VisioStdR_Retail"
                    sAcidList = "C5A7B998-FAE7-4F08-9802-9A2216F1EC2B;1B170697-99F8-4B3C-861D-FBDBE5303A8A;2F632ABD-D62A-46C8-BF95-70186096F1ED;980F9E3E-F5A8-41C8-8596-61404ADDF677;CEED49FA-52AD-4C90-B7D4-926ECDFD7F52;2E5F5521-6EFA-4E29-804D-993C1D1D7406;F97246FA-C8CE-4D41-B6F7-AE718E7891C5;E67ED831-B287-418B-ABC3-D68403A36166;DAE597CE-5823-4C77-9580-7268B93A4B23"
            End Select
        Case UCase("VisioStdVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioStdVL_KMS_Client;VisioStdVL_MAK"
                    sAcidList = "AA2A7821-1827-4C2C-8F1D-4513A34DDA97;44151C2D-C398-471F-946F-7660542E3369"
                Case 15
                    sSkuList  = "VisioStdVL_KMS_Client;VisioStdVL_MAK"
                    sAcidList = "AC4EFAF0-F81F-4F61-BDF7-EA32B02AB117;44A1F6FF-0876-4EDB-9169-DBB43101EE89"
            End Select
        Case UCase("VisioStdXVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "VisioStdXC2RVL_KMS_ClientC2R;VisioStdXC2RVL_MAKC2R"
                    sAcidList = "361FE620-64F4-41B5-BA77-84F8E079B1F7;1D1C6879-39A3-47A5-9A6D-ACEEFA6A289D"
            End Select
        Case UCase("WCServer")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "WCServer_ByPass;WCServer_BypassTrial"
                    sAcidList = "926E4E17-087B-47D1-8BD7-91A394BC6196;2D45B535-68E6-44E0-8496-FC74F1491566"
            End Select
        Case UCase("WCServerNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WCServer_Bypass"
                    sAcidList = "85A235B4-B6EC-4DAE-A3B3-AE6AF54F0040"
                Case 15
                    sSkuList  = "WCServer_Bypass"
                    sAcidList = "D6B57A0D-AE69-4A3E-B031-1F993EE52EDC"
            End Select
        Case UCase("WSS")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "WSS_ByPass"
                    sAcidList = "BEED1F75-C398-4447-AEF1-E66E1F0DF91E"
            End Select
        Case UCase("WSSLPK")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "WSSLPK_ByPass"
                    sAcidList = "D2C34FD2-557A-4A16-8E54-FABC7F720BEB"
            End Select
        Case UCase("WSSLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WSSLPK_Bypass"
                    sAcidList = "0594E026-263F-4C86-B364-F94AF2AB6874"
                Case 15
                    sSkuList  = "WSSLPK_Bypass"
                    sAcidList = "DAE4B8B2-832A-4726-968C-F81A2487E000"
            End Select
        Case UCase("WSSNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WSS_Bypass"
                    sAcidList = "A35DF3E6-6069-4D46-9E39-05ED7E8E2D7B"
                Case 15
                    sSkuList  = "WSS_Bypass"
                    sAcidList = "9FF54EBC-8C12-47D7-854F-3865D4BE8118"
            End Select
        Case UCase("WacServerLPK2019None")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WacServerLPK2019_Bypass"
                    sAcidList = "16BC5DDD-5B76-4A22-89CE-113D3BA5C852"
            End Select
        Case UCase("WacServerLPKNone")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WacServerLPK_Bypass"
                    sAcidList = "5B45E2C1-3D5A-4C24-AD16-B26499F64FC2"
                Case 15
                    sSkuList  = "WacServerLPK_Bypass"
                    sAcidList = "F3B8948B-6F0F-4297-B174-FE3E71416BCA"
            End Select
        Case UCase("Word")
            Select Case iVersionMajor
                Case 14
                    sSkuList  = "Word_KMS_Client;Word_MAK;Word_OEM_Perp;Word_Retail;Word_SubPrepid;Word_Trial"
                    sAcidList = "2D0882E7-A4E7-423B-8CCC-70D91E0158B1;98D4050E-9C98-49BF-9BE1-85E12EB3AB13;BED40A3E-6ACA-4512-8012-70AE831A2FC5;DB3BBC9C-CE52-41D1-A46F-1A1D68059119;99279F42-6DE2-4346-87B1-B0EC99C7525C;195E23D7-E0B7-4C30-8A30-8E9941AFD07E"
            End Select
        Case UCase("Word2019Retail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Word2019R_Grace;Word2019R_OEM_Perp;Word2019R_Retail;Word2019R_Trial"
                    sAcidList = "AFB40FCE-9886-413D-A6BE-9658C744AFB3;EF401233-F6E6-499E-A694-5C5B5D509D18;72CEE1C2-3376-4377-9F25-4024B6BAADF8;FB731819-2CAB-4AA8-9F51-3004DF6C1CFF"
            End Select
        Case UCase("Word2019Volume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "Word2019VL_KMS_Client_AE;Word2019VL_MAK_AE"
                    sAcidList = "059834FE-A8EA-4BFF-B67B-4D006B5447D3;FE5FE9D5-3B06-4015-AA35-B146F85C4709"
            End Select
        Case UCase("WordRetail")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WordR_Grace;WordR_OEM_Perp;WordR_Retail;WordR_Trial"
                    sAcidList = "6285E8B3-2D05-4954-BF85-65159DEEB966;60F6C5C4-96C1-4FFB-9719-388276C61605;CACAA1BF-DA53-4C3B-9700-11738EF1C2A5;E52D2A85-80E5-435B-882F-8884792D8ABF"
                Case 15
                    sSkuList  = "WordR_Grace;WordR_OEM_Perp;WordR_Retail;WordR_Trial"
                    sAcidList = "92485559-060B-44A7-9FC4-207C7A9BD39C;6A817637-8CA1-45D5-A53D-6B97FB8AF382;191509F2-6977-456F-AB30-CF0492B1E93A;AEA37447-9EE8-4EF5-8032-B0C955B6A4F5"
            End Select
        Case UCase("WordVolume")
            Select Case iVersionMajor
                Case 16
                    sSkuList  = "WordVL_KMS_Client;WordVL_MAK"
                    sAcidList = "BB11BADF-D8AA-470E-9311-20EAF80FE5CC;C3000759-551F-4F4A-BCAC-A4B42CBF1DE2"
                Case 15
                    sSkuList  = "WordVL_KMS_Client;WordVL_MAK"
                    sAcidList = "D9F5B1C6-5386-495A-88F9-9AD6B41AC9B3;9CEDEF15-BE37-4FF0-A08A-13A045540641"
            End Select
    End Select
End Sub 'MapConfigNameToLicense

'-------------------------------------------------------------------------------
'   GetLicTypeFromAcid
'
'   Obtain the license type channel linked to the ACID
'   Return the license type as string 
'   
'-------------------------------------------------------------------------------
Function GetLicTypeFromAcid (sAcid)
    On Error Resume Next
    Dim sReturn

    sReturn = ""

    Select Case sAcid
    Case "259DE5BE-492B-44B3-9D78-9645F848F7B0","0FBDE535-558A-4B6E-BDF7-ED6691AA7188","C5D855EE-F32B-4A1C-97A8-F0A28CE02F9C","B695D087-638E-439A-8BE3-57DB429D8A77","2CD676C7-3937-4AC7-B1B5-565EDA761F13","B7D84C2B-0754-49E4-B7BE-7EE321DCE0A9","12275A09-FEC0-45A5-8B59-446432F13CD6","3391E125-F6E4-4B1E-899C-A25E6092D40D","781038A1-800F-4DDF-AB26-70A98774F8AD","35466B1A-B17B-4DFB-A703-F74E2A1F5F5E","86E17AEA-A932-42C4-8651-95DE6CB37A08","7930E9BD-12B8-4B32-88CE-733B4A594EDF","E0132983-CA20-4390-8BBA-8FDA37E7C86B","F3B8948B-6F0F-4297-B174-FE3E71416BCA","D6B57A0D-AE69-4A3E-B031-1F993EE52EDC","9FF54EBC-8C12-47D7-854F-3865D4BE8118","DAE4B8B2-832A-4726-968C-F81A2487E000","68CBDF33-A8E4-4E51-B27D-1E777922395E","E4083CC2-9C14-4444-9428-99F4E1828C85","0B7586A3-C196-43B2-8195-A010CCF9A612","8662EA53-55DB-4A44-B72A-5E10822F9AAB","D98FA85C-4FD0-40D4-895C-6CC11DF2899F","6A7D5DA3-A27C-403C-87A0-BB4044A0EFBE","436366DE-5579-4F24-96DB-3893E4400030","D20B83AB-A959-4FA3-9C28-82C14D74285A","F3D55933-8E0C-4CD0-AC62-F98B3666C02D","AEF27390-E282-44EF-B18A-649AF2928DA9","AD96EEE1-FD16-4D77-8A5A-91B0C9A2369A","6061FDEF-2320-46DF-B9F0-33E63D003080","CD3F389B-747E-4730-A9BF-87349D3F737D","E7851824-6844-4D85-ABE7-7CA026FB3314","3F456727-422A-4134-8F62-A37380017039","D4639CC7-7B3A-4AC6-99C4-B0AF9ACE7B27","20AB14AF-1B6A-4C6B-A469-43EC0FE5AF78","5B45E2C1-3D5A-4C24-AD16-B26499F64FC2","16BC5DDD-5B76-4A22-89CE-113D3BA5C852","85A235B4-B6EC-4DAE-A3B3-AE6AF54F0040","A35DF3E6-6069-4D46-9E39-05ED7E8E2D7B","0594E026-263F-4C86-B364-F94AF2AB6874","745FB377-0A59-4CA9-B9A9-C359557A2C4E","71391382-EFB5-48B3-86BD-0450650FED7D","B5300EA3-8300-42D6-A3CD-4BFFA76397BA","60CD7BD4-F23C-4E58-9D81-968591E793EA","3FDFBCC8-B3E4-4482-91FA-122C6432805C","1F230F82-E3BA-4053-9B6D-E8F061A933FC","9C2C3D66-B43B-4C22-B105-67BDD5BB6E85","D5595F62-449B-4061-B0B2-0CBAD410BB51","C1CEDA8B-C578-4D5D-A4AA-23626BE4E234","F3329A70-BB26-4DD1-AF64-68E10C1AE635","288B2C51-35B5-41CA-AA5B-524DA113B4BD","ED21638F-97FF-4A65-AD9B-6889B93065E2","A4F824FC-C80E-4111-9884-DCEECE7D81A1","08460AA2-A176-442C-BDCA-26928704D80B","1328E89E-7EC8-4F7E-809E-7E945796E511","DAF85377-642C-4222-BCD8-DF5925C70AFB","B78DF69E-0966-40B1-AE85-30A5134DEDD0","2745E581-565A-4670-AE90-6BF7C57FFE43","4A8825CE-A527-4A42-B59C-17B22CFE644F","926E4E17-087B-47D1-8BD7-91A394BC6196","BEED1F75-C398-4447-AEF1-E66E1F0DF91E","D2C34FD2-557A-4A16-8E54-FABC7F720BEB","DE3337D8-860B-4E99-B54C-AC38836147CB","503EAC89-92BC-4134-988E-0F8658B40C5E","5B8E93B1-E8AE-414E-8927-180D295F8F0A","5B11FD8E-E03C-45DC-B35C-95145F4706EB"
        sReturn = "Bypass"
    Case "B13B737A-E990-4819-BC00-1A12D88963C3","7C7F5B37-6BBA-49D2-B0DE-27CC76BCE67D","2BEB303E-66C6-4422-B2EC-5AEA48B75EE5","0B1ACA01-5C25-468F-809D-DA81CB49AC3A","F10D4C70-F7CC-452A-B4B8-F12E3D6F4EEC","8FC4269F-A845-4D1F-9DF0-9F499C92D9CB","A42E403B-B141-4FA2-9781-647604D09BFB","8C893555-CB79-4BAA-94EA-780E01C3AB15","B2C0B444-3914-4ACB-A0B8-7CF50A8F7AA0","F48AC5F3-0E33-44D8-A6A9-A629355F3F0C","8F23C830-C19B-4249-A181-B20B6DCF7DBE","88BED06D-8C6B-4E62-AB01-546D6005FE97","3B02A9FF-CED3-46D6-9298-A4334829748F","4365667B-8304-463E-B542-2DF8D0A73EA9","2BCDDDBE-4EBE-4728-9594-625E26137761","84902853-59F6-4B20-BC7C-DE4F419FEFAD","BC4C1C97-9013-4033-A0DD-9DC9E6D6C887","2D45B535-68E6-44E0-8496-FC74F1491566","14DE70F9-E4DE-40DD-ABE8-65AC4ADD0C72","75FFD06C-AB06-476A-8BC3-2A31701B97EA","67ADC8CD-7DDB-4885-8504-C1328260CEBB"
        sReturn = "BypassTrial"
    Case "14D002D8-B1ED-4AB4-BE6C-5F3FA7ED2F75","AAA4ABAA-369E-4D7D-93E6-4C320A017A3C","A7693F85-CCA9-4878-9A32-8BE2773D9331","B31C9322-8FA5-454D-AC52-100D51EDFEE5","C7136135-DC7F-46C4-9818-EA246A597480"
        sReturn = "Bypass30"
    Case "8D071DB8-CDE7-4B90-8862-E2F6B54C91BF","6330BD12-06E5-478A-BDF4-0CD90C3FFE65","1C3432FE-153B-439B-82CC-FB46D6F5983A","CBF97833-C73A-4BAF-9ED3-D47B3CFF51BE","2F05F7AA-6FF7-49AB-884A-14696A7BD1A4","062FDB4F-3FCB-4A53-9B36-03262671B841","298A586A-E3C1-42F0-AFE0-4BCFDC2E7CD0","B845985D-10AE-449A-96AF-43CEDCE9363D","42C3EF3F-AB7C-482A-8060-B2E57B25D4BA","40E9B240-879B-4BA4-BFB0-2E94CA998EEB","BC7BAF08-4D97-462C-8411-341052402E71","82E42CB5-1741-46FB-8F2F-AC8414741E8D","6A9BD082-8C07-462C-8D19-18C2CAE0FB02","64F26796-F91A-43B1-87A3-3B2E45922D71","6F04E8C9-AD8A-4147-A647-DE0071BC703F","5F2C45C3-232D-43A1-AE90-40CD3462E9E4","E7B23B5D-BDB1-495C-B7CE-CC13652ACF2F","61E77797-5801-4361-9AC4-40DD2B07F60A","45E65598-7AE9-4943-8C4D-DE6D9A508108","BBFEDDD0-412D-49D5-BDA4-3F427736DE94","AE423093-CA28-4C04-A566-631C39C5D186","D72D694A-C20B-491A-9629-BEAFC48ECBD0","338E5858-1E3A-4ACD-900E-7DBEF543BB9B","DA892D82-124E-444B-BD01-B109155CB2AE","EB264A11-A2EE-4740-8AC2-3751C8859374","220842AE-9681-41F2-AF5A-4DBDE01FACC1","13104647-EE01-4016-AFAF-BBA564215D6E","19C742D3-FAAE-4A79-89B2-4F4E5C0F5B7C","CBCAF071-B911-40A1-A6C1-7BECCA481910","60AFA663-984D-47A6-AC9C-00346FF5E8F0","C726993A-CD6C-499F-B672-F36AECAFFB20","D0C2BC53-DB5F-4A9E-84AC-17C8862C27DA"
        sReturn = "BypassTrial180"
    Case "2AA09C36-3E86-4580-8E8D-E67DFC13A5A6","0C6911A0-1FD3-4425-A4B3-DA478D3BF807","81D507D6-E729-436D-AD4D-FED2CBE4E566","7AB38A1C-A803-4CA5-8EBB-F8786C4B4FAA","35EC6E0E-2DF4-4629-9EE3-D525E806B988","18C10CE7-C025-4615-993C-2E0C32F38EFA","8BF9E61F-3905-4E9B-99C4-1C3946C1FCA7"
        sReturn = "BypassTrial365"
    Case "6912CCDF-557A-497C-9903-3DE6CE9FA631"
        sReturn = "DeltaTrial"
    Case "80C94D2C-4DDE-47AE-82D2-A2ADDE81E653","360E9813-EA13-4152-B020-B1D0BBF1AC17","323277B1-D81D-4329-973E-497F413BC5D0","D95CFA27-EDDF-4116-8D28-BBDB1E07B0A5","0900883A-7F90-4A04-831D-69B5881A0C1C","CE6A8540-B478-4070-9ECB-2052DD288047","FBE35AC1-57A8-4D02-9A23-5F97003C37D3","0EA305CE-B708-4D79-8087-D636AB0F1A4D","E1F5F599-D875-48CA-93C4-E96C473B29E7","3D0631E3-1091-416D-92A5-42F84A86D868","D7279DD0-E175-49FE-A623-8FC2FC00AFC4","3AD61E22-E4FE-497F-BDB1-3E51BD872173","8DED1DA3-5206-4540-A862-E8473B65D742","FD97BCB1-8F3C-4185-91D2-A2AB7EE278C2","A3A4593B-97DC-4364-9910-70202AF2D0B5","34EBB57F-0773-4DF5-8EBD-BF55D194DC69","A75ADD9F-3982-464D-93FE-7D3BEC07FB46","F3A4939A-92A7-47CC-9D05-7C7DB72DD968","2A3D3010-CDAB-4F07-8BF2-0E3639BC6B98","A9419E0F-8A3F-4E58-A143-E4B4803F85D2","AE7B1E26-3AEE-4FE3-9C5B-88F05E36CD34","6883893B-9DD9-4943-ACEB-58327AFDC194","1B686580-9FB1-4B88-BFBA-EAE7C0DA31AD","DB8E8683-A848-473B-B2E7-D1DE4D042095","370155F7-D1EC-4B8B-9FBA-EE8ACC6E0BB7","024EA285-2685-48BC-87EF-79B48CC8C027","F97246FA-C8CE-4D41-B6F7-AE718E7891C5","92485559-060B-44A7-9FC4-207C7A9BD39C","195CBC52-71CC-4909-A7A3-90D745D26465","DFC79B08-1388-44E8-8CBE-4E4E42E9F73A","381823F8-7779-4908-AAF3-5929F6BE6FDB","FC61B360-59EE-41A2-8586-E66E7B43DB89","2C18EFFF-A05D-485C-BF41-E990736DB92E","D95CFA27-EDDF-4116-8D28-BBDB1E07B0A5","6379F5C8-BAA4-482E-A5CC-38ADE1A57348","FAFAAABC-4BC0-4EFE-A9FC-B8AA32D0E101","1ECBF9C7-7C43-483E-B2AD-06EF5C5744C4","A9833863-A86B-4DEA-A034-7FDEB7A6276B","8507F630-BEE7-4833-B054-C1CA1A43990C","3D0631E3-1091-416D-92A5-42F84A86D868","81E676D5-B92C-4253-AA47-36A0EA89F5ED","D7279DD0-E175-49FE-A623-8FC2FC00AFC4","3AD61E22-E4FE-497F-BDB1-3E51BD872173","8DED1DA3-5206-4540-A862-E8473B65D742","66C66DB4-55EA-427B-90A8-B850B2B87FD0","5464FE8F-BEFA-43A3-956B-3A1FFB59A280","6D47BDDA-2331-4956-A1D6-E57889F326AB","C860DD35-031F-49B8-B500-24C17A564291","34EBB57F-0773-4DF5-8EBD-BF55D194DC69","B444EF2B-340F-4E88-9E32-FCB747EE5D04","9C31B372-0652-41B9-BDCE-34EB16400E9F","E5101C42-817D-432D-9783-8C03134A51C2","EB972895-440E-4863-BE98-BA075B9721D0","2A3D3010-CDAB-4F07-8BF2-0E3639BC6B98","7A0560C5-21ED-4518-AD41-B7F870B9FD1A","7D351D13-E2F4-49EA-A63C-63A85B1D5630","CA5B3EEA-C055-4ACF-BC78-187DB21C7DB5","086C4557-2A7A-49B2-B4CE-3D2F93F8C8CA","587BECE0-14F5-4039-AA1F-7CFCE036FCCC","52C4D79F-6E1A-45B7-B479-36B666E0A2F8","70D9CEB6-6DFA-4DA4-B413-18C1C3C76E2E","A70DB9BD-3910-427E-B565-BB8DAAAAAFFF","6CB2AFA8-FFFF-49FE-91CF-5FFBB0C9FD2B","69632864-D608-472C-AD91-9B13E195C641","17C5A0EB-24EE-4EB3-91B6-6D9F9CB1ACF6","42419D2A-5756-47C5-9AD3-E1548E26A4EE","C7D96193-844A-44BD-93D0-76E4B785248E","DB2760F6-F36B-403C-A0E7-0B59CDE7427B","5821EC16-77A9-4404-99C8-2756DC6D4C3C","8497DB0B-415B-4900-A20E-ECEEC4A6D350","AD310F32-BD8D-40AD-A882-BC7FCBAD5259","AFB40FCE-9886-413D-A6BE-9658C744AFB3","6285E8B3-2D05-4954-BF85-65159DEEB966"
        sReturn = "Grace"
    Case "1DC00701-03AF-4680-B2AF-007FFC758A1F","E914EA6E-A5FA-4439-A394-A9BB3293CA09"
        sReturn = "KMS_Automation"
    Case "8CE7E872-188C-4B98-9D90-F8F90B7AAD02","CEE5D470-6E3B-4FCC-8C2B-D17428568A9F","8947D0B8-C33B-43E1-8C56-9B674C052832","CA6B6639-4AD6-40AE-A575-14DEE07F6430","09ED9640-F020-400A-ACD8-D7D867DFD9C2","EF3D4E49-A53D-4D81-A2B1-2CA6C2556B2C","AB586F5C-5256-4632-962F-FEFD8B49E6F4","ECB7C192-73AB-4DED-ACF4-2399B095D0CC","45593B1D-DFB1-4E91-BBFB-2D5D0CE2227A","DF133FF7-BF14-4F95-AFE3-7B48E7E331EF","5DC7BF61-5EC9-4996-9CCB-DF806A2D0EFE","6F327760-8C5C-417C-9B61-836A98287E0C","B50C4F75-599B-43E8-8DCD-1081A7967241","EA509E87-07A1-4A45-9EDC-EBA5A39F36AF","9DA2A678-FB6B-4E67-AB84-60DD6A9C819A","92236105-BB67-494F-94C7-7F7A607929BD","E558389C-83C3-4B29-ADFE-5E4D7F46C358","9ED833FF-4F92-4F36-B370-8683A4F13275","2D0882E7-A4E7-423B-8CCC-70D91E0158B1","6EE7622C-18D8-4005-9FB7-92DB644A279B","F7461D52-7C2B-43B2-8744-EA958E0BD09A","FB4875EC-0C6B-450F-B82B-AB57D8D1677F","A30B8040-D68A-423F-B0B5-9CE292EA5A8F","1B9F11E3-C85C-4E1B-BB29-879AD2C909E3","DC981C6B-FC8E-420F-AA43-F8F33E5C0923","EFE1F3E6-AEA2-4144-A208-32AA872B6545","771C3AFA-50C5-443F-B151-FF2546D863A0","8C762649-97D1-4953-AD27-B7E2C25B972E","4A5D124A-E620-44BA-B6FF-658961B33B9A","427A28D1-D17C-4ABF-B717-32C780BA6F07","B322DA9C-A2E2-4058-9E4E-F59A6970BD69","00C79FF1-6850-443D-BF61-71CDE0DE305F","B13AFB38-CD79-4AE5-9F7F-EED058D750CA","E13AC10E-75D0-4AFF-A0CD-764982CF541C","AC4EFAF0-F81F-4F61-BDF7-EA32B02AB117","D9F5B1C6-5386-495A-88F9-9AD6B41AC9B3","67C0FC0C-DEBA-401B-BF8B-9C8AD8395804","C3E65D36-141F-4D2F-A303-A842EE756A29","9CAABCCB-61B1-4B4B-8BEC-D10A3C3AC2CE","D8CACE59-33D2-4AC7-9B1B-9B72339C51C8","EC9D9265-9D1E-4ED0-838A-CDC20F2551A1","D70B1BBA-B893-4544-96E2-B7A318091C33","4F414197-0FC2-4C01-B68A-86CBB9AC254C","DA7DDABC-3FBE-4447-9E01-6AB7440B4CD4","D450596F-894D-49E0-966A-FD39ED4C4C64","041A06CB-C5B8-4772-809F-416D03D16654","83E04EE1-FA8D-436D-8994-D31A862CAB77","DEDFA23D-6ED1-45A6-85DC-63CAE0546DE6","6BF301C1-B94A-43E9-BA31-D494598C47FB","AA2A7821-1827-4C2C-8F1D-4513A34DDA97","BB11BADF-D8AA-470E-9311-20EAF80FE5CC"
        sReturn = "KMS_Client"
    Case "9E9BCEEB-E736-4F26-88DE-763F87DCC485","237854E9-79FC-4497-A0C1-A70969691C6B","C8F8A301-19F5-4132-96CE-2DE9D4ADBD33","3131FD61-5E4F-4308-8D6D-62BE1987C92C","2CA2BF3F-949E-446A-82C7-E25A15EC78C4","1777F0E3-7392-4198-97EA-8AE4DE6F6381","85DD8B5F-EAA4-4AF3-A628-CCE9E77C9A03","9D3E4CCA-E172-46F1-A2F4-1D2107051444","734C6C6E-B0BA-4298-A891-671772B2BD1B","6912A74B-A5FB-401A-BFDB-2E3AB46F4B02","5B5CF08F-B81A-431D-B080-3450D8620565","E06D7DF3-AAD0-419D-8DFB-0AC37E2BDF39","059834FE-A8EA-4BFF-B67B-4D006B5447D3"
        sReturn = "KMS_Client_AE"
    Case "FC7C4D0C-2E85-4BB9-AFD4-01ED1476B5E9","829B8110-0E6F-4349-BCA4-42803577788D","CBBACA45-556A-4416-AD03-BDA598EAA7C8","0BC88885-718C-491D-921F-6F214349E79C","500F6619-EF93-4B75-BCB4-82819998A3CA","B234ABE3-0857-4F9C-B05A-4DC314F85557","361FE620-64F4-41B5-BA77-84F8E079B1F7"
        sReturn = "KMS_ClientC2R"
    Case "BFE7A195-4F8F-4F0B-A622-CF13C7D16864","2E28138A-847F-42BC-9752-61B03FFF33CD","70512334-47B4-44DB-A233-BE5EA33B914C","98EBFE73-2084-4C97-932C-C0CD1643BEA7"
        sReturn = "KMS_Host"
    Case "95AB3EC8-4106-4F9D-B632-03C019D1D23F","71DC86FF-F056-40D0-8FFB-9592705C9B76","FDAD0DFA-417D-4B4F-93E4-64EA8867B7FD","85E22450-B741-430C-A172-A37962C938AF","533B656A-4425-480B-8E30-1A2358898350","6860B31F-6A67-48B8-84B9-E312B3485C4B","A9AEABD8-63B8-4079-A28E-F531807FD6B8","38252940-718C-4AA6-81A4-135398E53851","1CF57A59-C532-4E56-9A7D-FFA2FE94B474","11B39439-6B93-4642-9570-F2EB81BE2238","B6E9FAE1-1A0E-4C61-99D0-4AF068915378","FDF3ECB9-B56F-43B2-A9B8-1B48B6BAE1A7","191301D3-A579-428C-B0C7-D7988500F9E3","3D014759-B128-4466-9018-E80F6320D9D0","8090771E-D41A-4482-929E-DE87F1F47E46","1F76E346-E0BE-49BC-9954-70EC53A4FCFE","DD457678-5C3E-48E4-BC67-A89B7A3E3B44","36756CB8-8E69-4D11-9522-68899507CD6A","5980CF2B-E460-48AF-921E-0C2A79025D23","CAB3A4C4-F31A-4C12-AFA9-A0EECC86BD95","98D4050E-9C98-49BF-9BE1-85E12EB3AB13","4374022D-56B8-48C1-9BB7-D8F2FC726343","AC1AE7FD-B949-4E04-A330-849BC40638CF","4825AC28-CE41-45A7-9E6E-1FED74057601","9E016989-4007-42A6-8051-64EB97110CF2","E1264E10-AFAF-4439-A98B-256DF8BB156F","F33485A0-310B-4B72-9A0E-B1D605510DBD","B067E965-7521-455B-B9F7-C740204578A2","8D577C50-AE5E-47FD-A240-24986F73D503","E40DCB44-1D5C-4085-8E8F-943F33C4F004","ED34DC89-1C27-4ECD-8B2F-63D0F4CEDC32","2B9E4A37-6230-4B42-BEE2-E25CE86C8C7A","2B88C4F2-EA8F-43CD-805E-4D41346E18A7","38EA49F6-AD1D-43F1-9888-99A35D7C9409","A24CCA51-3D54-4C41-8A76-4031F5338CB2","3E4294DD-A765-49BC-8DBD-CF8B62A4BD3D","44A1F6FF-0876-4EDB-9169-DBB43101EE89","9CEDEF15-BE37-4FF0-A08A-13A045540641","3B2FA33F-CD5A-43A5-BD95-F49F3F546B0B","685062A7-6024-42E7-8C5F-6BB9E63E697F","2CD0EA7E-749F-4288-A05E-567C573B2A6C","23B672DA-A456-4860-A8F3-E062A501D7E8","50059979-AC6F-4458-9E79-710BCB41721A","9B4060C9-A7F5-4A66-B732-FAF248B7240F","82F502B5-B0B0-4349-BD2C-C560DF85B248","82E6B314-2A62-4E51-9220-61358DD230E6","C47456E3-265D-47B6-8CA0-C30ABBD0CA36","FCC1757B-5D5F-486A-87CF-C4D6DEDB6032","03CA3B9A-0869-4749-8988-3CBC9D9F51BB","0ED94AAC-2234-4309-BA29-74BDBB887083","295B2C03-4B1C-4221-B292-1411F468BD02","44151C2D-C398-471F-946F-7660542E3369","C3000759-551F-4F4A-BCAC-A4B42CBF1DE2"
        sReturn = "MAK"
    Case "385B91D6-9C2C-4A2E-86B5-F44D44A48C5F","05CB4E1D-CC81-45D5-A769-F34B09B9B391","92A99ED8-2923-4CB7-A4C5-31DA6B0B8CF3","13C2D7BF-F10D-42EB-9E93-ABF846785434","D4EBADD6-401B-40D5-ADF4-A5D4ACCD72D1","FDAA3C03-DC27-4A8D-8CBF-C3D843A28DDC","6755C7A7-4DFE-46F5-BCE8-427BE8E9DC62","40055495-BE00-444E-99CC-07446729B53E","15A430D4-5E3F-4E6D-8A0A-14BF3CAEE4C7","BEB5065C-1872-409E-94E2-403BCFB6A878","F41ABF81-F409-4B0D-889D-92B3E3D7D005","933ED0E3-747D-48B0-9C2C-7CEB4C7E473D","FE5FE9D5-3B06-4015-AA35-B146F85C4709"
        sReturn = "MAK_AE"
    Case "73FC2508-0FB1-475A-BD6F-2D42FD84C62A","16728639-A9AB-4994-B6D8-F81051E69833","431058F0-C059-44C5-B9E7-ED2DD46B6789","0B4E6CFA-3072-434C-B702-BD1836E6887C","7DF5867B-E38F-459F-A590-93C1DD2EF0C6","0594DC12-8444-4912-936A-747CA742DBDB","1D1C6879-39A3-47A5-9A6D-ACEEFA6A289D"
        sReturn = "MAKC2R"
    Case "1FDFB4E4-F9C9-41C4-B055-C80DAF00697D","EBEF9F05-5273-404A-9253-C5E252F50555","C86F17B4-67E0-459E-B785-F09FF54600E4","6303D14A-AFAD-431F-8434-81052A65F575","090896A0-EA98-48AC-B545-BA5DA0EB0C9C","215C841D-FFC1-4F03-BD11-5B27B6AB64CC","6BBE2077-01A4-4269-BF15-5BF4D8EFC0B2"
        sReturn = "OEM_ARM"
    Case "4B17D082-D27D-467D-9E70-40B805714C0A","EAEED721-9715-46FC-B2F8-03EEA2EF1FE2","209408DF-98DB-4EEB-B96B-D0B9A4B13468","F63B84D0-ED9D-4B05-99E4-19D33FD7AFBD","0EAAF923-70A2-48BD-A6F1-54CC1AA95C13","00B6BBFC-4091-4182-BB81-93A9A6DEB46A","5DBE2163-3FA9-464C-B8B7-CAADDE61E4FF","0E795CCE-5BAD-40B1-8803-CE71FB89031D","DDB12F7C-CE7E-4EE5-A01C-E6AF9EDBC020","8C7E3A91-E176-4AB3-84B9-A7E0EFB3A6DD","9F82274C-C0EF-4212-B8D9-97A6BFBC2DC7","95FF18B9-F2CF-4291-AB2E-BC17D54AA756","D79A3F4F-E768-4114-8D3A-7F9F45687F67","866AC003-01A0-49FD-A6EC-F8F56ABDCFAB","F6183964-39E3-46A9-A641-734041BAFF89","115A5CF2-D4CF-4627-91DC-839DF666D082","8C54246A-31FE-4274-90C7-687987735848","40EC9524-41D7-41D7-95C7-E5F185F92EA3","247E7706-0D68-4F56-8D78-2B8147A11CA8","1783C7A6-840C-4B33-AF05-2B1F5CD73527","7D4627B9-9467-4AA7-AE7F-892807D78D8F","7E05FC0C-7CE4-4849-BB0B-231BDF5DCA70","AE3ED6AE-2654-4B82-A4BA-331265BB8972","FABC9393-6174-4192-B3EE-6340E16CF90D","7B8EBE34-08FC-46C5-8BFA-161B12A43E41","8B1F0A02-07D6-411C-9FC7-9CAA3F86D1FE","4E6F61A8-989B-463C-9948-83B894540AD4","88B5EC99-C9D1-47F9-B1F2-3C6C63929B7B","BB42DD2B-070C-4F5B-947A-55F56A16D4F3","1359DCE0-0DC8-4171-8C43-BA8B9F2E5D1D","0B172E55-95AE-4C78-8C58-81AA98AB7F94","40BECF98-1D17-43EF-989F-1D92396A2741","BED40A3E-6ACA-4512-8012-70AE831A2FC5","70E3A52B-E7DA-4ABE-A834-B403A5776532","72D74828-C6B2-420B-AC78-68BF3A0E882C","C02FB62E-1CD5-4E18-BA25-E0480467FFAA","1E69B3EE-DA97-421F-BED5-ABCCE247D64E","BE465D55-C392-4D67-BF45-4DE7829EFC6E","6E2F71BC-1BA0-4620-872E-285F69C3141E","BA6BA8B7-2A1D-4659-BA45-4BAC007B1698","A1B2DD4A-29DD-4B74-A6FE-BF3463C00F70","5AAB8561-1686-43F7-9FF5-2C861DA58D17","1EFD9BEC-770F-4F5D-BC66-138ACAF5019E","AC867F78-3A34-4C70-8D80-53B6F8C4092C","4E26CAC1-E15A-4467-9069-CB47B67FE191","3178076B-5F4E-4ACE-A160-8AAE7F002944","3B5C7D36-1314-41C0-B405-15C558435F7D","519FBBDE-79A6-4793-8A84-57F6541579C9","BB00F022-69DA-4BE3-968D-D9BCD41CF814","CC802B96-22A4-4A90-8757-2ABB7B74484A","141F6878-2591-4231-A476-027432E28B2F","E67ED831-B287-418B-ABC3-D68403A36166","6A817637-8CA1-45D5-A53D-6B97FB8AF382","71BA37DB-38DE-4F2A-B290-1D031CE3A59E","DAF2E6D5-1424-47B9-8820-42B32E4BC3A2","74605A38-8EC0-482E-AC64-872E174E505B","740A90D3-A366-49FA-9845-1FA8257ED07B","1BA73CCC-3B8A-42B7-9936-0FB1863EA6EC","EB29320D-A377-4441-BE97-83C595A7139B","B464C2C6-4ADA-4931-9B62-B449136522D0","9A93B5EB-A24C-4178-8F91-F391F0BD3923","C02FB62E-1CD5-4E18-BA25-E0480467FFAA","BEA06487-C54E-436A-B535-58AEC833F05B","45C8342D-379D-4A4F-99C5-61141D3E0260","327DAB29-B51E-4A29-9204-D83A3B353D76","91AB0C97-4CBE-4889-ACA0-F0C9394046D7","480DF7D5-B237-45B2-96FB-E4624B12040A","8E2E934E-C319-4CCE-962E-F0546FB0AA8E","BF3AB6CF-1B0D-4250-BAD9-4B61ECA25316","F680A57C-24C4-4FDD-80A3-778877F63D0A","27D85D8B-FE2F-4515-9E37-A96056788046","6FAC4EBC-930E-4DF5-AAAE-9FF5A36A7C99","1810D5D6-2572-4ED1-BD96-08D77C288DCA","5AAB8561-1686-43F7-9FF5-2C861DA58D17","E66E9AA9-D6CE-43DA-8A34-EF62ECDAE720","BD6CCE28-11AF-484A-97FF-86422981B868","17A89B14-4244-4B97-BF6B-76D586CF91D0","65C693F4-3F1C-48A6-AFE6-A6EB78A65CCB","4E26CAC1-E15A-4467-9069-CB47B67FE191","54FF28C4-56AD-4168-AB75-7207F98DE5EB","BC39E263-86FB-4822-AD08-CB05F71CE855","2A9E12E5-48A2-497C-98D3-BE7170D86B13","DEEA27A6-E540-4B6B-9124-8A5CF73FD68B","6185F4C9-522C-4808-8508-946D28CEA707","447EFFDD-0073-4716-99F3-32EF2E155766","1640C37D-E2EB-4785-A70B-3C94A2E9B07E","6BD3E147-FF80-44FD-A3A6-B8E5E4FA9ADE","0E26561D-8B88-44CC-B87E-FAF56344E5F3","4A470B50-4D01-4FE5-910A-51BB0B7EB163","CDBD81EF-BB2C-41CC-BB3E-CA6366F3F176","6C1BED1D-0273-4045-90D2-E0836F3C380B","596BF8EC-7CAB-4A98-83AE-459DB70D24E4","26B394D7-7AD7-4AAB-8FCC-6EA678395A91","5829FD99-2B17-4BE4-9814-381145E49019","339A5901-9BDE-4F48-A88D-D048A42B54B1","AA64F755-8A7B-4519-BC32-CAB66DEB92CB","BDD04A4F-5804-4DC4-A1FF-8BC9FD0EC4EC","6BB0FF35-9392-4635-B209-642C331CF1EE","4EF70E76-5115-4BD2-9B12-3B0CB671AD37","ED967494-A18E-41A9-ABB8-9030C51BA685","C613537D-13A0-4D13-939C-BC876A2F5E8C","15265EDE-4272-4B92-9D47-EF584858D0AE","EF401233-F6E6-499E-A694-5C5B5D509D18","60F6C5C4-96C1-4FFB-9719-388276C61605"
        sReturn = "OEM_Perp"
    Case "E260D7A1-49DC-41EE-B275-33A544C96904","55671CEE-5384-454B-9160-975BD5441C4B","253CD5CD-F190-4A8C-80AC-0144DF8990F4","0C7137CB-463A-44CA-B6CD-F29EAE4A5201","EB778317-EA9A-4FBC-8351-F00EEC82E423","52263C53-5561-4FD5-A88D-2382945E5E8B","02BB3A66-D796-4938-94E3-1D6FF2A96B5B","E6930809-6796-466F-BC45-F0937362FDC1","6516C530-02EA-49D9-AD33-A824902B1065","AA1A1140-B8A1-4377-BFF0-4DCFF82F6ED1","5770F497-3F1D-4396-A729-18677AFE6359","003D2533-6A8C-4163-9A4E-76376B21ED15","0B40C0DA-6C2C-4BB6-9C05-EF8B30F96668","59C95D45-74A6-405E-815B-99CF2692577A","ADE257E8-4824-497F-B767-C0D06CDA3E4E","9CF873F3-5112-4C73-86C5-B94808F675E0","2D40ED06-955E-4725-80FC-850CDCBC1396","B688E64D-6305-4CCF-9981-CBF6823B3F65","65A4C65F-780B-4B57-B640-2AEA70761087","7BBE17F2-24C7-416A-8821-A7ACAF91A4AA","5D8AAAF2-3EAB-48E0-822A-6EFAE35980F7","FDA59085-BDB9-4ADC-B9D7-4B5E7DFD3EAC","AB2A78FF-5C52-494A-AE2A-99C0E481CA99","9540C768-CEDF-401C-B27A-15DD833E7053","949594B8-4FBF-44A7-BD8D-911E25BA2938","2A859A5B-3B1F-43EA-B3E3-C1531CC23363","5B7C4417-66C2-46DB-A2AF-DD1D811D8DCA","7E6F537B-AF16-43E6-B07F-F457A96F58FE","AE88E21A-D981-45A0-9813-3452F5390A1E","09322079-6043-4D33-9AB1-FFC268B8248E","E0C4D41F-B115-4D51-9E8C-63AF19B6EEB8","5917FA4B-7A37-4D70-B835-2755CF165E01","12F1FDF0-D744-48B6-9F9A-B3D7744BEEB5","A46815E5-C7E6-4F4F-8994-65BD19DC831D","D059DF6E-43A4-40EE-894D-5186F85AF9A0","E47D40AD-D9D0-4AB0-9C3D-B209E77E7AE8","58DB55C7-BB77-4C4D-A170-D7ED18A4AECE","B58FB371-5E42-46CE-927B-1A27AE89194C","D84CB3EA-CB53-466A-85C8-E8ED69D3F394","286E0A48-550E-4AFE-A223-27D510DE9E3B","A4CB27DA-2524-4998-92AE-6A836124BF9F","DAA7B9D0-F65A-482B-8AD3-474A33A05BE2","21757D20-006A-4BA0-8FDB-9DE272A11CF4","44938520-350C-4919-8E2F-E840C99AF4BF","F1997D24-27E6-4350-A9C7-50E63F85406A","7F3D4AF1-E298-4C30-8118-A0533B436C43","CB9FE322-23F2-4FC5-8569-0E63D0853072","F3449F6D-D512-4BBC-AC02-8C654845F6CB","436055A7-25A6-40FF-A124-B43A216FD369","3B694869-1556-4339-98C2-96C893040E31","7BAA6D3A-9EA9-47A8-8DFF-DC33F63CFEB3","11FC26E1-ADC8-4BEE-B42A-5C57D5FEC516","396D571C-B08F-4BEE-BE0B-28C5B640BC23","B6C0DE84-37D8-41B2-9EAF-8A80A008D58A","BCEF3379-4C3F-46D6-9922-8D3E513E75FA","C299D07D-803A-423C-88E9-029DE50D79ED","8C4F02D4-30C0-441A-860D-A68BE9EFC1EE","89B643DB-DFA5-4AF4-AEFC-43D17D98B084","D7E4AF12-6EB0-40DA-BF0E-D818BBC62AFB","183C0D5A-5AAE-45EB-B0CE-88009B8B3939","2E8E87E2-ED03-4DFB-8B40-0F73785B43FD","6B102F6C-50D4-4137-BF9F-5E446348188B","BE161169-49EC-4277-9802-6C5964D96593","D495DE19-448F-4C7E-8343-E54D7F336582","16C432FF-E13E-46A5-AB7D-A1CE7519119A","E0ECE06B-9123-4314-9570-338573410DE7","1752DF1B-8314-4790-9DF8-585D0BF1D0E3","3CF36478-D8B9-4FF4-9B42-0EDC60415E6A","CCB0F52E-312E-406E-A13A-566AF3C6B91D","634C8BC9-357C-4A70-A6B1-C2A9BC1BF9E5","B6512468-269D-4910-AD2D-A6A3D769FEBA","FEDA65D8-640B-4220-B8AE-6E221A990258","F337DD12-BF19-456B-82ED-B022B3032737","CB5E4CFF-1C96-43CA-91F8-0559DBA5456C","C2B1EC65-443F-4B2A-BF99-B46B244F0179","C9D3D19B-60A6-4213-BC72-C1A1A59DC45B","943CDA82-BDAC-4298-981F-21E97303DCCE","991A2521-DBB2-436C-82EB-C1BF93AC46FD","ED7482AE-C31E-4B62-A7CB-9291A0D616D2","63F119BD-171E-4034-B98F-A85759367616","792DE587-45FD-4F49-AD84-BFD40087792C","90984839-248D-4746-8F55-C943CAB009CF","62CCF9BD-1F60-414A-B766-79E98E55283B","F600E366-4C93-4CF0-BED5-BF2398759E00","174F331D-103E-4654-B6BD-C86D5B259A1A","DEA194EA-A396-4F91-88F2-77E501C8B9D2","B64CFCED-19A0-4C81-8B8D-A72DDC7488BA","C5B6E8E0-AD3D-4712-A5C6-FB7E7CC4DA59","A9AF7310-F6E0-40E2-B44A-076B2BA7E118","1F3BF441-3BD5-4DC5-B3DA-DB90450866F4","BDA57528-BB32-43DB-8CF0-36A32372F527","F98827AD-4B5F-48F1-B81B-FEE63DF5858E","B333EAB2-86CB-4B08-A5B8-A4C64AFC7A4B","F988DF4F-10DF-444F-909F-679BB9DF70D2","4D57374B-F5EA-4513-BB95-0456E424BCEE","9AF02C45-EE74-4400-A4E7-1251D11780A0","9459DA1D-F113-416B-AB41-A81D1C29A7A0","F8FF7F1C-1D40-4B84-9535-C51F7B3A70D0","04D62A0B-845D-48ED-A872-A064FCC56B07","5CA79892-8746-4F00-BB15-F3C799EDBCF4","E96A7A9C-162D-41E1-8C3D-D4C0929B6558","934ECDD6-7438-4464-B0FD-ED374AAB9CA6","5F1FB9D5-F241-4543-994F-36B37F888585","E5606A7D-8C29-4AD7-94B5-7427BF457616","4B9B31AA-00FA-481C-A25F-A2067F3FF24C","92B51A69-397D-4796-A5E8-D1672B349469","1AE119A6-40ED-4FA9-B52A-909841BD6E3F","FA25F074-80F9-44C1-B233-4019D3E70230","77A58083-E86D-45C1-B56E-77EC925167CF","523AABDF-9353-4151-9179-F708A2B48F6B","85133BEE-AD80-4CB4-977B-F16FD2DC3A20","3DC943FD-4E44-4FD6-A05F-8B421FF1C68A","8978F58A-BDC1-418C-AB3E-8017859FCC1E","9156E89A-1370-4079-AAA5-695DC9CEEAA2","B7CD5DC5-A2E6-4E74-BA44-F16B81B931F2","87D9FF03-50F4-4247-83E7-32ED575F2B1B","064A3A09-3B2B-441D-A86B-F3D47F2D2EC6","09775609-A49B-4F28-97AE-7F7740415C66","C9693FD7-4024-4544-BF5B-45160F7D5687","58E55634-63D9-434A-8B16-ADE5F84737B8","D545CCF1-CF46-41D0-8D55-D3E8B03CDB64","70A633E5-AB58-487A-B077-E410B30C5EA4","12367CA7-3B7D-44FD-8765-02469CA5004E","E260D7A1-49DC-41EE-B275-33A544C96904","55671CEE-5384-454B-9160-975BD5441C4B","253CD5CD-F190-4A8C-80AC-0144DF8990F4","0C7137CB-463A-44CA-B6CD-F29EAE4A5201","EB778317-EA9A-4FBC-8351-F00EEC82E423","52263C53-5561-4FD5-A88D-2382945E5E8B","02BB3A66-D796-4938-94E3-1D6FF2A96B5B","E6930809-6796-466F-BC45-F0937362FDC1","6516C530-02EA-49D9-AD33-A824902B1065","AA1A1140-B8A1-4377-BFF0-4DCFF82F6ED1","5770F497-3F1D-4396-A729-18677AFE6359","003D2533-6A8C-4163-9A4E-76376B21ED15","0B40C0DA-6C2C-4BB6-9C05-EF8B30F96668","59C95D45-74A6-405E-815B-99CF2692577A","ADE257E8-4824-497F-B767-C0D06CDA3E4E","9CF873F3-5112-4C73-86C5-B94808F675E0","2D40ED06-955E-4725-80FC-850CDCBC1396","B688E64D-6305-4CCF-9981-CBF6823B3F65","65A4C65F-780B-4B57-B640-2AEA70761087","7BBE17F2-24C7-416A-8821-A7ACAF91A4AA","5D8AAAF2-3EAB-48E0-822A-6EFAE35980F7","FDA59085-BDB9-4ADC-B9D7-4B5E7DFD3EAC","AB2A78FF-5C52-494A-AE2A-99C0E481CA99","9540C768-CEDF-401C-B27A-15DD833E7053","949594B8-4FBF-44A7-BD8D-911E25BA2938","2A859A5B-3B1F-43EA-B3E3-C1531CC23363","5B7C4417-66C2-46DB-A2AF-DD1D811D8DCA","7E6F537B-AF16-43E6-B07F-F457A96F58FE","AE88E21A-D981-45A0-9813-3452F5390A1E","09322079-6043-4D33-9AB1-FFC268B8248E","E0C4D41F-B115-4D51-9E8C-63AF19B6EEB8","5917FA4B-7A37-4D70-B835-2755CF165E01","12F1FDF0-D744-48B6-9F9A-B3D7744BEEB5","A46815E5-C7E6-4F4F-8994-65BD19DC831D","D059DF6E-43A4-40EE-894D-5186F85AF9A0","E47D40AD-D9D0-4AB0-9C3D-B209E77E7AE8","58DB55C7-BB77-4C4D-A170-D7ED18A4AECE","B58FB371-5E42-46CE-927B-1A27AE89194C","D84CB3EA-CB53-466A-85C8-E8ED69D3F394","286E0A48-550E-4AFE-A223-27D510DE9E3B","A4CB27DA-2524-4998-92AE-6A836124BF9F","DAA7B9D0-F65A-482B-8AD3-474A33A05BE2","21757D20-006A-4BA0-8FDB-9DE272A11CF4","44938520-350C-4919-8E2F-E840C99AF4BF","F1997D24-27E6-4350-A9C7-50E63F85406A","7F3D4AF1-E298-4C30-8118-A0533B436C43","CB9FE322-23F2-4FC5-8569-0E63D0853072","F3449F6D-D512-4BBC-AC02-8C654845F6CB","436055A7-25A6-40FF-A124-B43A216FD369","3B694869-1556-4339-98C2-96C893040E31","7BAA6D3A-9EA9-47A8-8DFF-DC33F63CFEB3","11FC26E1-ADC8-4BEE-B42A-5C57D5FEC516","396D571C-B08F-4BEE-BE0B-28C5B640BC23","B6C0DE84-37D8-41B2-9EAF-8A80A008D58A","BCEF3379-4C3F-46D6-9922-8D3E513E75FA","C299D07D-803A-423C-88E9-029DE50D79ED","8C4F02D4-30C0-441A-860D-A68BE9EFC1EE","89B643DB-DFA5-4AF4-AEFC-43D17D98B084","D7E4AF12-6EB0-40DA-BF0E-D818BBC62AFB","183C0D5A-5AAE-45EB-B0CE-88009B8B3939","2E8E87E2-ED03-4DFB-8B40-0F73785B43FD","6B102F6C-50D4-4137-BF9F-5E446348188B","BE161169-49EC-4277-9802-6C5964D96593","D495DE19-448F-4C7E-8343-E54D7F336582","16C432FF-E13E-46A5-AB7D-A1CE7519119A","E0ECE06B-9123-4314-9570-338573410DE7","1752DF1B-8314-4790-9DF8-585D0BF1D0E3","3CF36478-D8B9-4FF4-9B42-0EDC60415E6A","CCB0F52E-312E-406E-A13A-566AF3C6B91D","634C8BC9-357C-4A70-A6B1-C2A9BC1BF9E5","B6512468-269D-4910-AD2D-A6A3D769FEBA","FEDA65D8-640B-4220-B8AE-6E221A990258","F337DD12-BF19-456B-82ED-B022B3032737","CB5E4CFF-1C96-43CA-91F8-0559DBA5456C","C2B1EC65-443F-4B2A-BF99-B46B244F0179","C9D3D19B-60A6-4213-BC72-C1A1A59DC45B","943CDA82-BDAC-4298-981F-21E97303DCCE","991A2521-DBB2-436C-82EB-C1BF93AC46FD","ED7482AE-C31E-4B62-A7CB-9291A0D616D2","63F119BD-171E-4034-B98F-A85759367616","792DE587-45FD-4F49-AD84-BFD40087792C","90984839-248D-4746-8F55-C943CAB009CF","62CCF9BD-1F60-414A-B766-79E98E55283B","F600E366-4C93-4CF0-BED5-BF2398759E00","174F331D-103E-4654-B6BD-C86D5B259A1A","DEA194EA-A396-4F91-88F2-77E501C8B9D2","B64CFCED-19A0-4C81-8B8D-A72DDC7488BA","C5B6E8E0-AD3D-4712-A5C6-FB7E7CC4DA59","A9AF7310-F6E0-40E2-B44A-076B2BA7E118","1F3BF441-3BD5-4DC5-B3DA-DB90450866F4","BDA57528-BB32-43DB-8CF0-36A32372F527","F98827AD-4B5F-48F1-B81B-FEE63DF5858E","B333EAB2-86CB-4B08-A5B8-A4C64AFC7A4B","F988DF4F-10DF-444F-909F-679BB9DF70D2","4D57374B-F5EA-4513-BB95-0456E424BCEE","9AF02C45-EE74-4400-A4E7-1251D11780A0","9459DA1D-F113-416B-AB41-A81D1C29A7A0","F8FF7F1C-1D40-4B84-9535-C51F7B3A70D0","04D62A0B-845D-48ED-A872-A064FCC56B07","5CA79892-8746-4F00-BB15-F3C799EDBCF4","E96A7A9C-162D-41E1-8C3D-D4C0929B6558","934ECDD6-7438-4464-B0FD-ED374AAB9CA6","5F1FB9D5-F241-4543-994F-36B37F888585","E5606A7D-8C29-4AD7-94B5-7427BF457616","4B9B31AA-00FA-481C-A25F-A2067F3FF24C","92B51A69-397D-4796-A5E8-D1672B349469","1AE119A6-40ED-4FA9-B52A-909841BD6E3F","FA25F074-80F9-44C1-B233-4019D3E70230","77A58083-E86D-45C1-B56E-77EC925167CF","523AABDF-9353-4151-9179-F708A2B48F6B","85133BEE-AD80-4CB4-977B-F16FD2DC3A20","3DC943FD-4E44-4FD6-A05F-8B421FF1C68A","8978F58A-BDC1-418C-AB3E-8017859FCC1E","9156E89A-1370-4079-AAA5-695DC9CEEAA2","B7CD5DC5-A2E6-4E74-BA44-F16B81B931F2","87D9FF03-50F4-4247-83E7-32ED575F2B1B","064A3A09-3B2B-441D-A86B-F3D47F2D2EC6","09775609-A49B-4F28-97AE-7F7740415C66","C9693FD7-4024-4544-BF5B-45160F7D5687","58E55634-63D9-434A-8B16-ADE5F84737B8","D545CCF1-CF46-41D0-8D55-D3E8B03CDB64","70A633E5-AB58-487A-B077-E410B30C5EA4","12367CA7-3B7D-44FD-8765-02469CA5004E"
        sReturn = "PIN"
    Case "9103F3CE-1084-447A-827E-D6097F68C895","FF693BF4-0276-4DDB-BB42-74EF1A0C9F4D","BA3E3833-6A7E-445A-89D0-7802A9A68588","22E6B96C-1011-4CD5-8B35-3C8FB6366B86","9D9FAF9E-D345-4B49-AFCE-68CB0A539C7C","72621D09-F833-4CC5-812B-2929A06DCF4F","A1BFBAE1-3092-4B33-BF77-0681BE6807FE","C3B2612D-58EE-4D9B-AD9D-7568DECB29F7","F88CFDEC-94CE-4463-A969-037BE92BC0E7","971CD368-F2E1-49C1-AEDD-330909CE18B6","9103F3CE-1084-447A-827E-D6097F68C895","2EA0D3C6-B498-4503-97A1-19AC88AE44E5"
        sReturn = "PrepidBypass"
    Case "4D463C2C-0505-4626-8CDB-A4DA82E2D8ED","4EAFF0D0-C6CB-4187-94F3-C7656D49A0AA","7004B7F0-6407-4F45-8EAC-966E5F868BDE","7B7D1F17-FDCB-4820-9789-9BEC6E377821","00495466-527F-442F-A681-F36FAD813F86","4EFBD4C4-5422-434C-8C25-75DA21B9381C","09E2D37E-474B-4121-8626-58AD9BE5776F","1CAEF4EC-ADEC-4236-A835-882F5AFD4BF0","7B0FF49B-22DA-4C74-876F-B039616D9A4E","C3AE020C-5A71-4CC5-A27A-2A97C2D46860","25FE4611-B44D-49CC-AE87-2143D299194E","D652AD8D-DA5C-4358-B928-7FB1B4DE7A7C","A963D7AE-7A88-41A7-94DA-8BB5635A8AF9","EF1DA464-01C8-43A6-91AF-E4E5713744F9","14F5946A-DEBC-4716-BABC-7E2C240FEC08","4671E3DF-D7B8-4502-971F-C1F86139A29A","09AE19CA-75F8-407F-BF70-8D9F24C8D222","3F7AA693-9A7E-44FC-9309-BB3D8E604925","FBF4AC36-31C8-4340-8666-79873129CF40","ACB51361-C0DB-4895-9497-1831C41F31A6","45AB9E9F-C1DC-4BC7-8E37-89E0C1241776","133C8359-4E93-4241-8118-30BB18737EA0","8B559C37-0117-413E-921B-B853AEB6E210","50AC2361-FE88-4E5E-B0B2-13ACC96CA9AE","A7971F62-61D0-4C67-ABCC-085C10CF470F","C4109E90-6C4A-44F6-B380-EF6137122F16","725714D7-D58F-4D12-9FA8-35873C6F7215","AA188B61-D3D3-443C-9DEC-5B42393EE5CB","47A5840C-8124-4A1F-A447-50168CD6833D","688F6589-2BD9-424E-A152-B13F36AA6DE1","71AF7E84-93E6-4363-9B69-699E04E74071","46C84AAD-65C7-482D-B82A-1EDC52E6989A","75BB133B-F5DD-423C-8321-3BD0B50322A5","42CBF3F6-4D5E-49C6-991A-0D99B8429A6D","8C5EDB5D-9AA0-47A7-9416-D61C7419A60A","98677603-A668-4FA4-9980-3F1F05F78F69","DBE3AEE0-5183-4FF7-8142-66050173CB01","AF2AFE5B-55DD-4252-AF42-E6F79CC07EBC","D3422CFB-8D8B-4EAD-99F9-EAB0CCD990D7","B6D2565C-341D-4768-AD7D-ADDBE00BB5CE","66CAD568-C2DC-459D-93EC-2F3CB967EE34","D0A97E12-08A1-4A45-ADD5-1155B204E766","0993043D-664F-4B2E-A7F1-FD92091FA81F","15A9D881-3184-45E0-B407-466A68A691B1","BA24D057-8B5F-462E-87FE-485038C68954","DB3BBC9C-CE52-41D1-A46F-1A1D68059119","AB4D047B-97CF-4126-A69F-34DF08E2F254","1B1D9BD5-12EA-4063-964C-16E7E87D6E08","CFAF5356-49E3-48A8-AB3C-E729AB791250","CD256150-A898-441F-AAC0-9F8F33390E45","98685D21-78BD-4C62-BC4F-653344A63035","44984381-406E-4A35-B1C3-E54F499556E2","FADA6658-BFC6-4C4E-825A-59A89822CDA8","3169C8DF-F659-4F95-9CC6-3115E6596E83","BEFEE371-A2F5-4648-85DB-A2C55FDF324C","0C4E5E7A-B436-4776-BB89-88E4B14687E2","7A75647F-636F-4607-8E54-E1B7D1AD8930","8B524BCC-67EA-4876-A509-45E46F6347E8","12004B48-E6C8-4FFA-AD5A-AC8D4467765A","17E9DF2D-ED91-4382-904B-4FED6A12CAF0","31743B82-BFBC-44B6-AA12-85D42E644D5B","44BC70E2-FB83-4B09-9082-E5557E0C2EDE","A1D1EFF9-1301-4709-BC2C-B6FCE5C158D8","F2435DE4-5FC0-4E5B-AC97-34F515EC5EE7","5517E6A2-739B-4822-946F-7F0F1C5934B1","41499869-4103-4D3B-9DA6-D07DF41B6E39","064383FA-1538-491C-859B-0ECAB169A0AB","C3A0814A-70A4-471F-AF37-2313A6331111","A884AB66-DE2E-4505-BA8B-3FB9DAF149ED","32255C0A-16B4-4CE2-B388-8A4267E219EB","A491EFAD-26CE-4ED7-B722-4569BA761D14","15D12AD4-622D-4257-976C-5EB3282FB93D","DAE597CE-5823-4C77-9580-7268B93A4B23","191509F2-6977-456F-AB30-CF0492B1E93A","518687BD-DC55-45B9-8FA6-F918E1082E83","BFA358B0-98F1-4125-842E-585FA13032E6","C201C2B7-02A1-41A8-B496-37C72910CD4A","424D52FF-7AD2-4BC7-8AC6-748D767B455D","7FE09EEF-5EED-4733-9A60-D7019DF11CAC","86834D00-7896-4A38-8FAE-32F20B86FA2B","AC9C8FB4-387F-4492-A919-D56577AD9861","522F8458-7D49-42FF-A93B-670F6B6176AE","4539AA2C-5C31-4D47-9139-543A868E5741","C28ACDB8-D8B3-4199-BAA4-024D09E97C99","E2127526-B60C-43E0-BED1-3C9DC3D5A468","B21367DF-9545-4F02-9F24-240691DA0E58","83AC4DD9-1B93-40ED-AA55-EDE25BB6AF38","20E359D5-927F-47C0-8A27-38ADBDD27124","5A670809-0983-4C2D-8AAD-D3C2C5B7D5D1","2747B731-0F1F-413E-A92D-386EC1277DD8","A9F645A1-0D6A-4978-926A-ABCB363B72A6","7E63CC20-BA37-42A1-822D-D5F29F33A108","F32D1284-0792-49DA-9AC6-DEB2BC9C80B6","1717C1E0-47D3-4899-A6D3-1022DB7415E0","D64EDC00-7453-4301-8428-197343FAFB16","499E61C7-900A-4882-B003-E8E106BDCB95","0D270EF7-5AAF-4370-A372-BC806B96ADB7","2CB19A15-BAB2-4FCB-ACEE-4BDE5BE207A5","0F42F316-00B1-48C5-ADA4-2F52B5720AD0","067DABCC-B31D-4350-96A6-4FCAD79CFBB7","BB7FFE5F-DAF9-4B79-B107-453E1C8427B5","E9F0B3FC-962F-4944-AD06-05C10B6BCD5E","5F472F1E-EB0A-4170-98E2-FB9E7F6FF535","A3072B8F-ADCC-4E75-8D62-FDEB9BDFAE57","84832881-46EF-4124-8ABC-EB493CDCF78E","DE52BD50-9564-4ADC-8FCB-A345C17F84F9","F053A7C7-F342-4AB8-9526-A1D6E5105823","6E0C1D99-C72E-4968-BCB7-AB79E03E201E","B639E55C-8F3E-47FE-9761-26C6A786AD6B","418D2B9F-B491-4D7F-84F1-49E27CC66597","1C9D26C1-4001-41BB-8B90-B83677CCC80D","FDFA34DD-A472-4B85-BEE6-CF07BF0AAA1C","2395220C-BE32-4E64-81F2-6D0BDDB03EB1","4A31C291-3A12-4C64-B8AB-CD79212BE45E","0758C976-EA87-4128-8C1C-F7AF1BAB76B9","A6F69D68-5590-4E02-80B9-E7233DFF204E","3846A94F-3005-401E-B663-AF1CF99813D5","2DFE2075-2D04-4E43-816A-EB60BBB77574","1F7413CF-88D2-4F49-A283-E4D757D179DF","4A582021-18C2-489F-9B3D-5186DE48F1CD","C76DBCBC-D71B-4F45-B5B3-B7494CB4E23E","72CEE1C2-3376-4377-9F25-4024B6BAADF8","CACAA1BF-DA53-4C3B-9700-11738EF1C2A5"
        sReturn = "Retail"
    Case "B3DADE99-DE64-4CEC-BCC7-A584D510782A","0BE50797-9053-4F15-B9B1-2F2C5A310816","84155B23-BAE0-4748-967B-40E12917B0BB","B21DA2D5-50F1-4C5C-BF59-07BAA35E25BA","19316117-30A8-4773-8FD9-7F7231F4E060","AFCA9E83-152D-48A8-A492-6D552E40EE8A","C7DF3516-425A-4A84-9420-0112F3094D90","D82665D5-2D8F-46BA-ABEC-FDF06206B956","31E631F4-EE62-4B1F-AEB6-3B30393E0045","C735DCC2-F5E9-4077-A72F-4B6D254DDC43","0209AC7B-8A4B-450B-92F2-B583152A2613","3D9F0DCA-BCF1-4A47-8F90-EB53C6111AFD","698FA94F-EB99-43BE-AB8C-5A085C36936C","8FB0D83E-2BCC-43CD-871A-6AD7A06349F4","148CE971-9CCE-4DE0-9FD6-4871BEBD5666","1EEA4120-6699-47E9-9E5D-2305EE108BAC","71FB05B7-19E2-4567-AF77-8F31681D39D2","9A13EB9C-006F-450A-9F59-4CEC1EAB88F5","4D06F72E-FD50-4BC2-A24B-D448D7F17EF2","F4C9C7E4-8C96-4513-ADA3-0A514B3AC5CF","28FE27A7-2E11-4C05-8DD0-E1F1C08DC3AE","E98EF0C0-71C4-42CE-8305-287D8721E26C","17B7CE1A-92A9-4B59-A7D0-E872D8A9A994","08CEF85D-8704-417E-A749-B87C7D218CAD","BC8275B7-D67A-4390-8C5E-CC57CFC74328","59EC6B79-F6F5-4ADD-A5A0-B755BFB77422","3513C04B-9085-43A9-8F9A-639993C19E80","0EC894E8-A5A9-48DE-9463-061C4801EE8F","4CC91C85-44A8-4834-B15D-FFEA4616E4E4","99279F42-6DE2-4346-87B1-B0EC99C7525C"
        sReturn = "SubPrepid"
    Case "2637E47C-CD16-45A1-8FF7-B7938723FD10","EFF11B33-79B0-4D87-B05F-AE5E4EC5F209","2C971DC3-E122-4DBD-8ADF-B5777799799A","3047CDE0-03E2-4BAE-ABC9-40AD640B418D","AE28E0AB-590F-4BE3-B7F6-438DDA6C0B1C","8D1E5912-B904-40A6-ADDD-8C7482879E87","EFF34BA1-2794-4908-9501-5190C8F2025E","A2B90E7A-A797-4713-AF90-F0BECF52A1DD","F2DE350D-3028-410A-BFAE-283E00B44D0E","69EC9152-153B-471A-BF35-77EC88683EAE","6337137E-7C07-4197-8986-BECE6A76FC33","537EA5B5-7D50-4876-BD38-A53A77CACA32","52287AA3-9945-44D1-9046-7A3373666821","C85FE025-B033-4F41-936B-B121521BA1C1","B58A5943-16EA-420F-A611-7B230ACD762C","28F64A3F-CF84-46DA-B1E8-2DFA7750E491","149DBCE7-A48E-44DB-8364-A53386CD4580","E3DACC06-3BC2-4E13-8E59-8E05F3232325","A8119E32-B17C-4BD3-8950-7D1853F4B412","6E5DB8A5-78E6-4953-B793-7422351AFE88","FF02E86C-FEF0-4063-B39F-74275CDDD7C3","BACD4614-5BEF-4A5E-BAFC-DE4C788037A2","0BC1DAE4-6158-4A1C-A893-807665B934B2","65C607D5-E542-4F09-AD0B-40D6A88B2702","4CB3D290-D527-45C2-B079-26842762FDD3","881D148A-610E-4F0A-8985-3BE4C0DB2B09","4031E9F3-E451-4E18-BEDD-91CAA3690C91","2F72340C-B555-418D-8B46-355944FE66B8","053D3F49-B913-4B33-935E-F930DECD8709","58D95B09-6AF6-453D-A976-8EF0AE0316B1","4B2E77A0-8321-42A0-B36D-66A2CDB27AF4","A56A3B37-3A35-4BBB-A036-EEE5F1898EEE","C5A7B998-FAE7-4F08-9802-9A2216F1EC2B","980F9E3E-F5A8-41C8-8596-61404ADDF677","69EC9152-153B-471A-BF35-77EC88683EAE","37256767-8FC0-4B6A-8B92-BC785BFD9D05","6337137E-7C07-4197-8986-BECE6A76FC33","2F5C71B4-5B7A-4005-BB68-F9FAC26F2EA3","537EA5B5-7D50-4876-BD38-A53A77CACA32","52287AA3-9945-44D1-9046-7A3373666821","C85FE025-B033-4F41-936B-B121521BA1C1","B58A5943-16EA-420F-A611-7B230ACD762C","28F64A3F-CF84-46DA-B1E8-2DFA7750E491","7984D9ED-81F9-4D50-913D-317ECD863065","CBECB6F5-DA49-4029-BE25-5945AC9750B3","149DBCE7-A48E-44DB-8364-A53386CD4580","E3DACC06-3BC2-4E13-8E59-8E05F3232325","A8119E32-B17C-4BD3-8950-7D1853F4B412","6E5DB8A5-78E6-4953-B793-7422351AFE88","FF02E86C-FEF0-4063-B39F-74275CDDD7C3","BACD4614-5BEF-4A5E-BAFC-DE4C788037A2","0BC1DAE4-6158-4A1C-A893-807665B934B2","65C607D5-E542-4F09-AD0B-40D6A88B2702","4CB3D290-D527-45C2-B079-26842762FDD3","881D148A-610E-4F0A-8985-3BE4C0DB2B09","4031E9F3-E451-4E18-BEDD-91CAA3690C91","2F72340C-B555-418D-8B46-355944FE66B8","053D3F49-B913-4B33-935E-F930DECD8709","58D95B09-6AF6-453D-A976-8EF0AE0316B1","4B2E77A0-8321-42A0-B36D-66A2CDB27AF4","A56A3B37-3A35-4BBB-A036-EEE5F1898EEE","C5A7B998-FAE7-4F08-9802-9A2216F1EC2B","980F9E3E-F5A8-41C8-8596-61404ADDF677"
        sReturn = "Subscription"
    Case "F5BEB18A-6861-4625-A369-9C0A2A5F512F","F3F68D9F-81E5-4C7D-9910-0FD67DC2DDF2","9E7FE6CF-58C6-4F36-9A1B-9C0BFC447ED5","742178ED-6B28-42DD-B3D7-B7C0EA78741B","A96F8DAE-DA54-4FAD-BDC6-108DA592707A","77F47589-2212-4E3B-AD27-A900CE813837","96353198-443E-479D-9E80-9A6D72FD1A99","B4AF11BE-5F94-4D8F-9844-CE0D5E0D8680","4C1D1588-3088-4796-B3B2-8935F3A0D886","AE1310F8-2F53-4994-93A3-C61502E91D04","539165C6-09E3-4F4B-9C29-EEC86FDF545F","594E9433-6492-434D-BFCE-4014F55D3909","5DF83BED-7E8E-4A28-80EF-D8B0A004CF3E","6C44E4BD-C6DB-474A-877E-1BB899ACA206","27B162B5-F5D2-40E9-8AF3-D42FF4572BD4","1B170697-99F8-4B3C-861D-FBDBE5303A8A","CEED49FA-52AD-4C90-B7D4-926ECDFD7F52","9E7FE6CF-58C6-4F36-9A1B-9C0BFC447ED5","CEBCF005-6D81-4C9E-8626-5D01C48B834A","742178ED-6B28-42DD-B3D7-B7C0EA78741B","A96F8DAE-DA54-4FAD-BDC6-108DA592707A","77F47589-2212-4E3B-AD27-A900CE813837","96353198-443E-479D-9E80-9A6D72FD1A99","B4AF11BE-5F94-4D8F-9844-CE0D5E0D8680","4C1D1588-3088-4796-B3B2-8935F3A0D886","AE1310F8-2F53-4994-93A3-C61502E91D04","539165C6-09E3-4F4B-9C29-EEC86FDF545F","594E9433-6492-434D-BFCE-4014F55D3909","5DF83BED-7E8E-4A28-80EF-D8B0A004CF3E","6C44E4BD-C6DB-474A-877E-1BB899ACA206","27B162B5-F5D2-40E9-8AF3-D42FF4572BD4","1B170697-99F8-4B3C-861D-FBDBE5303A8A","CEED49FA-52AD-4C90-B7D4-926ECDFD7F52"
        sReturn = "SubTest"
    Case "92847EEE-6935-4585-817D-14DCFFE6F607","D46BBFF5-987C-4DAC-9A8D-069BAC23AE2A","26421C15-1B53-4E53-B303-F320474618E7","812837CB-040B-4829-8E22-15C4919FEBA4","951C0FE6-40A8-4400-8003-EEC0686FFBC4","FB3540BC-A824-4C79-83DA-6F6BD3AC6CCB","86AF9220-9D6F-4575-BEF0-619A9A6AA005","AE9D158F-9450-4A0C-8C80-DD973C58357C","95820F84-6E8C-4D22-B2CE-54953E9911BC","46D2C0BD-F912-4DDC-8E67-B90EADC3F83C","26B6A7CE-B174-40AA-A114-316AA56BA9FC","B6B47040-B38E-4BE2-BF6A-DABF0C41540A","E538D623-C066-433D-A6B7-E0708B1FADF7","DFC5A8B0-E9FD-43F7-B4CA-D63F1E749711","6AE65B85-E04E-4368-80A7-786D5766325E","2030A84D-D6C7-4968-8CEC-CF4737ACC337","2F756D47-E1B1-4D2A-922B-4D76F35D007D","C5A706AC-9733-4061-8242-28EB639F1B29","090506FC-50F8-4C00-B8C7-91982A2A7C99","2B456DC8-145A-4D2D-9C72-703D1CD8C50E","330A4ACE-9CC1-4AF5-8D36-8D0681194618","B480F090-28FE-4A67-8885-62322037A0CD","75F08E14-5CF5-4D59-9CEF-DA3194B6FD24","79EEBE53-530A-4869-B56B-B4FFE0A1B831","402DFB6D-631E-4A3F-9A5A-BE753BFD84C5","2F632ABD-D62A-46C8-BF95-70186096F1ED","2E5F5521-6EFA-4E29-804D-993C1D1D7406","26421C15-1B53-4E53-B303-F320474618E7","EE6555BB-7A24-4287-99A0-5B8B83C5D602","812837CB-040B-4829-8E22-15C4919FEBA4","B2F7AD5B-0E90-4AB7-9DA5-D829A6EA3F1E","951C0FE6-40A8-4400-8003-EEC0686FFBC4","FB3540BC-A824-4C79-83DA-6F6BD3AC6CCB","86AF9220-9D6F-4575-BEF0-619A9A6AA005","AE9D158F-9450-4A0C-8C80-DD973C58357C","95820F84-6E8C-4D22-B2CE-54953E9911BC","24FC428E-A37E-4996-AC66-8EA0304A152D","B27B3D00-9A95-4FCD-A0C2-118CBD5E699B","46D2C0BD-F912-4DDC-8E67-B90EADC3F83C","26B6A7CE-B174-40AA-A114-316AA56BA9FC","B6B47040-B38E-4BE2-BF6A-DABF0C41540A","E538D623-C066-433D-A6B7-E0708B1FADF7","DFC5A8B0-E9FD-43F7-B4CA-D63F1E749711","6AE65B85-E04E-4368-80A7-786D5766325E","2030A84D-D6C7-4968-8CEC-CF4737ACC337","2F756D47-E1B1-4D2A-922B-4D76F35D007D","C5A706AC-9733-4061-8242-28EB639F1B29","090506FC-50F8-4C00-B8C7-91982A2A7C99","2B456DC8-145A-4D2D-9C72-703D1CD8C50E","330A4ACE-9CC1-4AF5-8D36-8D0681194618","B480F090-28FE-4A67-8885-62322037A0CD","75F08E14-5CF5-4D59-9CEF-DA3194B6FD24","79EEBE53-530A-4869-B56B-B4FFE0A1B831","402DFB6D-631E-4A3F-9A5A-BE753BFD84C5","2F632ABD-D62A-46C8-BF95-70186096F1ED","2E5F5521-6EFA-4E29-804D-993C1D1D7406"
        sReturn = "SubTrial"
    Case "ED826596-52C4-4A81-85BD-0F343CBC6D67","8FC26056-52E4-4519-889F-CDBEDEAC7C31","DD1E0912-6816-4DC2-A8BF-AA2971DB0E25","4790B2A5-BBF2-4C26-976F-D7736E516CCE","1DFBB6C1-0C4D-44E9-A0EA-77F59146E011","3850C794-B06F-4633-B02F-8AC4DF0A059F","F4F25E2B-C13C-4256-8E4C-5D4D82E1B862","B49D9ABE-7F30-40AA-9A4C-BDE08A14832D","131E900A-EFA8-412E-BA20-CB0F4BE43054","533D80CB-BF68-48DB-AB3E-165B5377599E","4F0B7650-A09D-4180-976D-76D8B31EA1B4","8A730535-A343-44EB-91A2-B1650DE0A872","EFEEFEDE-44B4-44FF-84D5-01B67A380024","AB324F86-606E-42EC-93E3-FC755989FE19","8CC3794C-4B71-44EA-BAAE-D95CC1D17042","2C8ACFCA-F0D9-4CCF-BA28-2F2C47DA8BA5","20F986B2-048B-4810-9421-73F31CA97657","1C57AD8F-60BE-4CE0-82EC-8F55AA09751F","DF01848D-8F9D-4589-9198-4AC51F2547F3","42122F59-2850-485E-B0C0-1AACA1C88923","23037F94-D654-4F38-962F-FF5B15348630","694E35B9-F965-47D7-AA19-AB2783224ADF","FD2BBCED-F8DB-45BC-B4D6-AC05A47D3691","F510F8DE-4325-4461-BD33-571EDBE0A933","8C5FA740-5DCA-43F9-BE1B-D0281BCF9779","1A069855-55EC-4CCF-9C45-2AC5D500F792","4519ABCF-23DB-487B-AC28-7C9EBE801716","7616C283-5C5B-4054-B52C-902F03E4DCDF","673EA9EA-9BC0-463F-93E5-F77655E46630","A27DF0C4-AE71-4DDD-BBEB-6D6222FE3A17","195E23D7-E0B7-4C30-8A30-8E9941AFD07E","E2E45C14-401F-4955-B05C-B6BDA44BF303","5290A34F-2DE8-4965-B53A-2D2976CB7B35","7868FC3D-ABA3-4B5D-B4EA-419C608DAB45","BB8DF749-885C-47D8-B33A-7E5A402EF4A3","6CA43757-B2AF-4C42-9E28-F763C023EDC8","EAB6228E-54F9-4DDE-A218-291EA5B16C0A","6A88EBA4-8B91-4553-8764-61129E1E01C7","2DA73A73-4C36-467F-8A2D-B49090D983B7","B6DDD089-E96F-43F7-9A77-440E6F3CD38E","AF3A9181-8B97-4B28-B2B1-A7AC6F8C3A05","621292D6-D737-40F8-A6E8-A9D0B852AEC2","52C5FBED-8BFB-4CB4-8BFF-02929CD31A98","1FCA7624-3A67-4A02-9EF3-6C363E35C8CD","41937580-5DDD-4806-9089-5266D567219F","DB56DEC3-34F2-4BC5-A7B9-ECC3CC51C12A","615FEE6D-2E96-487F-98B2-51A892788A70","603EDDF5-E827-4233-AFDE-4084B3F3B30C","F35E39C1-A41F-47C9-A204-2CA3C4B13548","AEA37447-9EE8-4EF5-8032-B0C955B6A4F5","35CB0CD2-0B1C-47CD-B365-6E85FC2F755A","BB6FE552-10FB-4C0E-A4BE-6B867E075735","04CAFB94-2615-419B-825A-F259E8A19D2B","A49503A7-3084-4140-B087-6B30B258F92E","C08057D3-830A-43A6-8009-4E2D018EC257","3A4A2CF6-5892-484D-B3D5-FE9CC5A20D78","3E0F37CC-584B-46B1-873E-C44AD57D6605","4EBC82D3-A2FA-49F3-9159-70E8DF6D45D5","861343B3-3190-494C-A76E-A9215C7AE254","99D89E49-3F21-4D6E-BB17-3E4EE4FE8726","8BDA5AA2-F2DB-4124-A513-7BE79DA4E499","D229947A-7C85-4D12-944B-7536B5B05B13","04388DC4-ABCA-4924-B238-AC4273F8F74B","EC9BF3B9-335A-4D5B-B9C7-7C3937CFA1F0","2DCD3F0B-AAFB-48D7-9F5F-01E14160AC43","04EC8535-E9EE-4778-A3F0-098CBF761E1E","FBBC4AAB-247A-43DB-8CE6-69ABE89A82E5","D26F7324-6471-444B-A8F5-25BBB7CB2DF4","E48D4EBB-FD30-43B2-BE87-0EB2335EC2C9","EDC83D99-C10D-4ADF-8D0C-98342381820B","39A1BE8C-9E7F-4A75-81F4-21CFAC7CBECB","F6C61FE0-58F0-49E0-9D79-12F5130B48AC","AEEDF8F7-8832-41B1-A9C8-13F2991A371C","E18C439E-424A-4E83-8F1B-5F9DD021B38D","83BD3F0A-A827-46C9-8B30-3F73BF933CFE","C8CE6ADC-EDE7-4CE2-8E7B-C49F462AB8C3","E1FEF7E5-6886-458C-8E45-7C1E9DAAB00C","771BE4FF-9148-4299-86AD-1A3EC0C10B10","FBD9E09F-6B81-4FFB-97EF-E8919A0C9E06","90BFBEDA-4F37-4D5E-8DA9-5EF0BDBC500F","BDB32189-C636-46CD-A011-FFB834B552D9","4CA376C1-6DEE-4AB4-B84D-0866CDD292CF","6CE96173-4198-49FC-9924-CC8EEC1A7285","D51BB17C-7E2B-45F4-A9F3-D2D444CC1E13","A17F9ED0-C3D4-4873-B3B8-D7E049B459EC","FB731819-2CAB-4AA8-9F51-3004DF6C1CFF","E52D2A85-80E5-435B-882F-8884792D8ABF"
        sReturn = "Trial"
    Case "762A768D-5882-4A57-A889-C387E85C10B0"
        sReturn = "ZeroGrace"
    End Select

    GetLicTypeFromAcid = sReturn
End Function

'-------------------------------------------------------------------------------
'   GetProductReleaseIdFromPrimaryProductId
'
'   Map the primary product id to the ProductReleaseId (PRId)
'   Return a string with possible PRId's 
'   
'-------------------------------------------------------------------------------

Function GetProductReleaseIdFromPrimaryProductId (iVersionMajor, sSkuId)
    On Error Resume Next
    Dim sPRId

    Select Case sSkuId
        Case "000F"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "MondoRetail;MondoVolume"
                Case 15
                    sPRId  = "MondoRetail;MondoVolume"
                Case 14
                    sPRId  = "Mondo"
            End Select
        Case "0011"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "KMSHostVolume;ProPlusRetail;ProPlusVolume"
                Case 15
                    sPRId  = "KMSHostVolume;ProPlusRetail;ProPlusVolume"
                Case 14
                    sPRId  = "ProPlus;ProPlusAcad;ProPlusMSDN;Sub4"
            End Select
        Case "0012"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "O365BusinessRetail;StandardRetail;StandardVolume"
                Case 15
                    sPRId  = "StandardRetail;StandardVolume"
                Case 14
                    sPRId  = "Standard;StandardAcad;StandardMSDN"
            End Select
        Case "0013"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeBusinessRetail;HomeBusinessPipcRetail"
                Case 15
                    sPRId  = "HomeBusinessRetail"
                Case 14
                    sPRId  = "HomeBusiness;HomeBusinessDemo"
            End Select
        Case "0014"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProfessionalRetail"
                Case 15
                    sPRId  = "ProfessionalRetail"
                Case 14
                    sPRId  = "Professional;ProfessionalAcad;ProfessionalDemo"
            End Select
        Case "0015"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "AccessRetail;AccessVolume"
                Case 15
                    sPRId  = "AccessRetail;AccessVolume"
                Case 14
                    sPRId  = "Access"
            End Select
        Case "0016"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ExcelRetail;ExcelVolume"
                Case 15
                    sPRId  = "ExcelRetail;ExcelVolume"
                Case 14
                    sPRId  = "Excel"
            End Select
        Case "0017"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "SPDRetail"
                Case 14
                    sPRId  = "SPD"
            End Select
        Case "0018"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PowerPointRetail;PowerPointVolume"
                Case 15
                    sPRId  = "PowerPointRetail;PowerPointVolume"
                Case 14
                    sPRId  = "PowerPoint"
            End Select
        Case "0019"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PublisherRetail;PublisherVolume"
                Case 15
                    sPRId  = "PublisherRetail;PublisherVolume"
                Case 14
                    sPRId  = "Publisher"
            End Select
        Case "001A"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "OutlookRetail;OutlookVolume"
                Case 15
                    sPRId  = "OutlookRetail;OutlookVolume"
                Case 14
                    sPRId  = "Outlook"
            End Select
        Case "001B"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "WordRetail;WordVolume"
                Case 15
                    sPRId  = "WordRetail;WordVolume"
                Case 14
                    sPRId  = "Word"
            End Select
        Case "001C"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "AccessRuntimeRetail"
                Case 15
                    sPRId  = "AccessRuntimeRetail"
                Case 14
                    sPRId  = "AccessRuntime"
            End Select
        Case "0029"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "HSExcel"
            End Select
        Case "002B"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "HSWord"
            End Select
        Case "002F"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeStudentRetail;HomeStudentVNextRetail"
                Case 15
                    sPRId  = "HomeStudentRetail"
                Case 14
                    sPRId  = "HomeStudent;HomeStudentDemo"
            End Select
        Case "0033"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PersonalRetail;PersonalPipcRetail"
                Case 15
                    sPRId  = "PersonalRetail"
                Case 14
                    sPRId  = "Personal;PersonalDemo;PersonalPrepaid"
            End Select
        Case "0037"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "HSPowerPoint"
            End Select
        Case "003A"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectStdRetail;ProjectStdVolume;ProjectStdXVolume"
                Case 15
                    sPRId  = "ProjectStdRetail;ProjectStdVolume"
                Case 14
                    sPRId  = "ProjectStd"
            End Select
        Case "003B"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectProRetail;ProjectProVolume;ProjectProXVolume"
                Case 15
                    sPRId  = "ProjectProRetail;ProjectProVolume"
                Case 14
                    sPRId  = "ProjectPro;ProjectProMSDN"
            End Select
        Case "003D"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "OEM"
            End Select
        Case "0044"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "InfoPathRetail;InfoPathVolume"
                Case 14
                    sPRId  = "InfoPath"
            End Select
        Case "004B"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PTKNone"
                Case 15
                    sPRId  = "PTKNone"
                Case 14
                    sPRId  = "PTK"
            End Select
        Case "0051"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioProRetail;VisioProVolume;VisioProXVolume"
                Case 15
                    sPRId  = "VisioProRetail;VisioProVolume"
                Case 14
                    sPRId  = "VisioPro;VisioProMSDN"
            End Select
        Case "0053"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioStdRetail;VisioStdVolume;VisioStdXVolume"
                Case 15
                    sPRId  = "VisioStdRetail;VisioStdVolume"
                Case 14
                    sPRId  = "VisioStd"
            End Select
        Case "0055"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioLPKNone;VisioLPK2019None"
                Case 15
                    sPRId  = "VisioLPKNone"
                Case 14
                    sPRId  = "VisioLPK"
            End Select
        Case "0058"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "VisioPrem"
            End Select
        Case "0059"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioStdRetail"
                Case 15
                    sPRId  = "VisioStdRetail"
            End Select
        Case "005A"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioProRetail"
                Case 15
                    sPRId  = "VisioProRetail"
            End Select
        Case "0066"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "Starter;StarterPrem"
            End Select
        Case "008B"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "SmallBusBasics;SmallBusBasicsMSDN"
            End Select
        Case "00A1"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "OneNoteRetail;OneNoteVolume"
                Case 15
                    sPRId  = "OneNoteFreeRetail;OneNoteRetail;OneNoteVolume"
                Case 14
                    sPRId  = "OneNote"
            End Select
        Case "00A3"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "OneNoteFreeRetail"
                Case 14
                    sPRId  = "HSOneNote"
            End Select
        Case "00B3"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectProRetail"
                Case 15
                    sPRId  = "ProjectProRetail"
            End Select
        Case "00B5"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectLPKNone;ProjectLPK2019None"
                Case 15
                    sPRId  = "ProjectLPKNone"
                Case 14
                    sPRId  = "ProjectLPK"
            End Select
        Case "00B6"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectStdRetail"
                Case 15
                    sPRId  = "ProjectStdRetail"
            End Select
        Case "00BA"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "GrooveRetail;GrooveVolume"
                Case 14
                    sPRId  = "Groove"
            End Select
        Case "00CE"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeStudentARM2019Retail;HomeStudentARMRetail"
                Case 15
                    sPRId  = "HomeStudentARMRetail"
            End Select
        Case "00D4"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "O365ProPlusRetail"
                Case 15
                    sPRId  = "O365ProPlusRetail"
            End Select
        Case "00D5"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "O365ProPlusRetail;O365SmallBusPremRetail"
                Case 15
                    sPRId  = "O365SmallBusPremRetail"
            End Select
        Case "00D6"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "O365EduCloudRetail;O365HomePremRetail"
                Case 15
                    sPRId  = "O365HomePremRetail"
            End Select
        Case "00DA"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeStudentPlusARM2019Retail;HomeStudentPlusARMRetail"
                Case 15
                    sPRId  = "HomeStudentPlusARMRetail"
            End Select
        Case "00E6"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PersonalPipcRetail"
                Case 15
                    sPRId  = "PersonalPipcRetail"
            End Select
        Case "00E7"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeBusinessPipcRetail"
                Case 15
                    sPRId  = "HomeBusinessPipcRetail"
            End Select
        Case "00E8"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProfessionalPipcRetail"
                Case 15
                    sPRId  = "ProfessionalPipcRetail"
            End Select
        Case "00E9"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "O365BusinessRetail"
                Case 15
                    sPRId  = "O365BusinessRetail"
            End Select
        Case "00EA"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "LyncAcademicRetail"
            End Select
        Case "0100"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "OfficeLPKNone;OfficeLPK2019None"
                Case 15
                    sPRId  = "OfficeLPKNone"
                Case 14
                    sPRId  = "OfficeLPK"
            End Select
        Case "011D"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "ProPlusSub"
            End Select
        Case "011E"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "HomeBusinessSub"
            End Select
        Case "011F"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "ProjectProSub"
            End Select
        Case "012C"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusinessRetail;SkypeforBusinessVolume;SkypeServiceBypassRetail"
                Case 15
                    sPRId  = "LyncRetail;LyncVolume"
            End Select
        Case "012D"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusinessEntryRetail"
                Case 15
                    sPRId  = "LyncEntryRetail"
            End Select
        Case "012E"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusinessVDINone"
                Case 15
                    sPRId  = "LyncVDINone"
            End Select
        Case "1017"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "GrooveServerEMS"
            End Select
        Case "1018"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "GrooveServer"
            End Select
        Case "110D"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "MOSSNone;MOSSFISEntNone;MOSSFISStdNone;MOSSPremiumNone"
                Case 14
                    sPRId  = "MOSS;MOSSFISEnt;MOSSFISStd;MOSSPremium"
            End Select
        Case "110F"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectServerNone"
                Case 15
                    sPRId  = "ProjectServerNone"
                Case 14
                    sPRId  = "ProjectServer"
            End Select
        Case "1110"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "WSSNone"
                Case 15
                    sPRId  = "WSSNone"
                Case 14
                    sPRId  = "WSS"
            End Select
        Case "1112"
            Select Case iVersionMajor
                Case 15
                    sPRId  = "ServerLPKNone"
                Case 14
                    sPRId  = "ServerLPK"
            End Select
        Case "1119"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "WSSLPKNone"
                Case 15
                    sPRId  = "WSSLPKNone"
                Case 14
                    sPRId  = "WSSLPK"
            End Select
        Case "111C"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "SearchServer"
            End Select
        Case "112D"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "WCServerNone"
                Case 15
                    sPRId  = "WCServerNone"
                Case 14
                    sPRId  = "WCServer"
            End Select
        Case "1137"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "SearchServerExpress"
            End Select
        Case "1139"
            Select Case iVersionMajor
                Case 14
                    sPRId  = "FastServer;FastServerIA;FastServerIB"
            End Select
        Case "1157"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "WacServerLPKNone;WacServerLPK2019None"
                Case 15
                    sPRId  = "WacServerLPKNone"
            End Select
        Case "1167"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "MOSSNone;MOSSFISEntNone;MOSSFISStdNone;MOSSPremiumNone"
            End Select
        Case "1168"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ServerLPKNone"
            End Select
        Case "4000"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Professional2019Retail"
            End Select
        Case "4001"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectPro2019Retail;ProjectPro2019Volume"
            End Select
        Case "4002"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProPlus2019Retail;ProPlus2019Volume"
            End Select
        Case "4003"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioPro2019Retail;VisioPro2019Volume"
            End Select
        Case "4011"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "KMSHost2019Volume"
            End Select
        Case "4012"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Standard2019Retail;Standard2019Volume"
            End Select
        Case "4013"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeBusiness2019Retail"
            End Select
        Case "4015"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Access2019Retail;Access2019Volume"
            End Select
        Case "4016"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Excel2019Retail;Excel2019Volume"
            End Select
        Case "4018"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PowerPoint2019Retail;PowerPoint2019Volume"
            End Select
        Case "4019"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Publisher2019Retail;Publisher2019Volume"
            End Select
        Case "401A"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Outlook2019Retail;Outlook2019Volume"
            End Select
        Case "401B"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Word2019Retail;Word2019Volume"
            End Select
        Case "401C"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "AccessRuntime2019Retail"
            End Select
        Case "402F"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "HomeStudent2019Retail"
            End Select
        Case "4033"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "Personal2019Retail"
            End Select
        Case "403A"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "ProjectStd2019Retail;ProjectStd2019Volume"
            End Select
        Case "404B"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "PTK2019None"
            End Select
        Case "4053"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "VisioStd2019Retail;VisioStd2019Volume"
            End Select
        Case "412C"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusiness2019Retail;SkypeforBusiness2019Volume"
            End Select
        Case "412D"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusinessEntry2019Retail"
            End Select
        Case "412E"
            Select Case iVersionMajor
                Case 16
                    sPRId  = "SkypeforBusinessVDI2019None"
            End Select
    End Select
    GetProductReleaseIdFromPrimaryProductId = sPRId
End Function

'=======================================================================================================

Function GetLicErrDesc(hErr)
    On Error Resume Next
    Select Case "0x"& hErr
    Case "0x0" : GetLicErrDesc = "Success."
    Case "0xC004B001" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B002" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B003" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B004" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B005" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B006" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B007" : GetLicErrDesc = "The activation server reported that the computer could not connect to the activation server."
    Case "0xC004B008" : GetLicErrDesc = "The activation server determined that the computer could not be activated."
    Case "0xC004B009" : GetLicErrDesc = "The activation server determined that the license is invalid."
    Case "0xC004B011" : GetLicErrDesc = "The activation server determined that your computer clock time is not correct. You must correct your clock before you can activate."
    Case "0xC004B100" : GetLicErrDesc = "The activation server determined that the computer could not be activated."
    Case "0xC004C001" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C002" : GetLicErrDesc = "The activation server determined there is a problem with the specified product key."
    Case "0xC004C003" : GetLicErrDesc = "The activation server determined the specified product key has been blocked."
    Case "0xC004C004" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C005" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C006" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C007" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C008" : GetLicErrDesc = "The activation server determined that the specified product key could not be used."
    Case "0xC004C009" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C00A" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C00B" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C00C" : GetLicErrDesc = "The activation server experienced an error."
    Case "0xC004C00D" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C00E" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C00F" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C010" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C011" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C012" : GetLicErrDesc = "The activation server experienced a network error."
    Case "0xC004C013" : GetLicErrDesc = "The activation server experienced an error."
    Case "0xC004C014" : GetLicErrDesc = "The activation server experienced an error."
    Case "0xC004C020" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key has exceeded its limit."
    Case "0xC004C021" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key extension limit has been exceeded."
    Case "0xC004C022" : GetLicErrDesc = "The activation server reported that the re-issuance limit was not found."
    Case "0xC004C023" : GetLicErrDesc = "The activation server reported that the override request was not found."
    Case "0xC004C016" : GetLicErrDesc = "The activation server reported that the specified product key cannot be used for online activation."
    Case "0xC004C017" : GetLicErrDesc = "The activation server determined the specified product key has been blocked for this geographic location."
    Case "0xC004C015" : GetLicErrDesc = "The activation server experienced an error."
    Case "0xC004C050" : GetLicErrDesc = "The activation server experienced a general error."
    Case "0xC004C030" : GetLicErrDesc = "The activation server reported that time based activation attempted before start date."
    Case "0xC004C031" : GetLicErrDesc = "The activation server reported that time based activation attempted after end date."
    Case "0xC004C032" : GetLicErrDesc = "The activation server reported that new time based activation not available."
    Case "0xC004C033" : GetLicErrDesc = "The activation server reported that time based product key not configured for activation."
    Case "0xC004C04F" : GetLicErrDesc = "The activation server reported that no business rules available to activate specified product key."
    Case "0xC004C700" : GetLicErrDesc = "The activation server reported that business rule cound not find required input."
    Case "0xC004C750" : GetLicErrDesc = "The activation server reported that NULL value specified for business property name and Id."
    Case "0xC004C751" : GetLicErrDesc = "The activation server reported that property name specifies unknown property."
    Case "0xC004C752" : GetLicErrDesc = "The activation server reported that property Id specifies unknown property."
    Case "0xC004C755" : GetLicErrDesc = "The activation server reported that it failed to update product key binding."
    Case "0xC004C756" : GetLicErrDesc = "The activation server reported that it failed to insert product key binding."
    Case "0xC004C757" : GetLicErrDesc = "The activation server reported that it failed to delete product key binding."
    Case "0xC004C758" : GetLicErrDesc = "The activation server reported that it failed to process input XML for product key bindings."
    Case "0xC004C75A" : GetLicErrDesc = "The activation server reported that it failed to insert product key property."
    Case "0xC004C75B" : GetLicErrDesc = "The activation server reported that it failed to update product key property."
    Case "0xC004C75C" : GetLicErrDesc = "The activation server reported that it failed to delete product key property."
    Case "0xC004C764" : GetLicErrDesc = "The activation server reported that the product key type is unknown."
    Case "0xC004C770" : GetLicErrDesc = "The activation server reported that the product key type is being used by another user."
    Case "0xC004C780" : GetLicErrDesc = "The activation server reported that it failed to insert product key record."
    Case "0xC004C781" : GetLicErrDesc = "The activation server reported that it failed to update product key record."
    Case "0xC004C401" : GetLicErrDesc = "The Vista Genuine Advantage Service determined that the installation is not genuine."
    Case "0xC004C600" : GetLicErrDesc = "The Vista Genuine Advantage Service determined that the installation is not genuine."
    Case "0xC004C801" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C802" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C803" : GetLicErrDesc = "The activation server determined the specified product key has been revoked."
    Case "0xC004C804" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C805" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C810" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C811" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C812" : GetLicErrDesc = "The activation server determined that the specified product key has exceeded its activation count."
    Case "0xC004C813" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C814" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0xC004C815" : GetLicErrDesc = "The activation server determined the license is invalid."
    Case "0xC004C816" : GetLicErrDesc = "The activation server reported that the specified product key cannot be used for online activation."
    Case "0xC004E001" : GetLicErrDesc = "The Software Licensing Service determined that the specified context is invalid."
    Case "0xC004E002" : GetLicErrDesc = "The Software Licensing Service reported that the license store contains inconsistent data."
    Case "0xC004E003" : GetLicErrDesc = "The Software Licensing Service reported that license evaluation failed."
    Case "0xC004E004" : GetLicErrDesc = "The Software Licensing Service reported that the license has not been evaluated."
    Case "0xC004E005" : GetLicErrDesc = "The Software Licensing Service reported that the license is not activated."
    Case "0xC004E006" : GetLicErrDesc = "The Software Licensing Service reported that the license contains invalid data."
    Case "0xC004E007" : GetLicErrDesc = "The Software Licensing Service reported that the license store does not contain the requested license."
    Case "0xC004E008" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
    Case "0xC004E009" : GetLicErrDesc = "The Software Licensing Service reported that the license store is not initialized."
    Case "0xC004E00A" : GetLicErrDesc = "The Software Licensing Service reported that the license store is already initialized."
    Case "0xC004E00B" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
    Case "0xC004E00C" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be opened or created."
    Case "0xC004E00D" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be written."
    Case "0xC004E00E" : GetLicErrDesc = "The Software Licensing Service reported that the license store could not read the license file."
    Case "0xC004E00F" : GetLicErrDesc = "The Software Licensing Service reported that the license property is corrupted."
    Case "0xC004E010" : GetLicErrDesc = "The Software Licensing Service reported that the license property is missing."
    Case "0xC004E011" : GetLicErrDesc = "The Software Licensing Service reported that the license store contains an invalid license file."
    Case "0xC004E012" : GetLicErrDesc = "The Software Licensing Service reported that the license store failed to start synchronization properly."
    Case "0xC004E013" : GetLicErrDesc = "The Software Licensing Service reported that the license store failed to synchronize properly."
    Case "0xC004E014" : GetLicErrDesc = "The Software Licensing Service reported that the license property is invalid."
    Case "0xC004E015" : GetLicErrDesc = "The Software Licensing Service reported that license consumption failed."
    Case "0xC004E016" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
    Case "0xC004E017" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
    Case "0xC004E018" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
    Case "0xC004E019" : GetLicErrDesc = "The Software Licensing Service determined that validation of the specified product key failed."
    Case "0xC004E01A" : GetLicErrDesc = "The Software Licensing Service reported that invalid add-on information was found."
    Case "0xC004E01B" : GetLicErrDesc = "The Software Licensing Service reported that not all hardware information could be collected."
    Case "0xC004E01C" : GetLicErrDesc = "This evaluation product key is no longer valid."
    Case "0xC004E01D" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (CD-AB)"
    Case "0xC004E01E" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (AB-AB)"
    Case "0xC004E01F" : GetLicErrDesc = "The new product key cannot be used on this installation of Windows. Type a different product key. (AB-CD)"
    Case "0xC004E020" : GetLicErrDesc = "The Software Licensing Service reported that there is a mismatched between a policy value and information stored in the OtherInfo section."
    Case "0xC004E021" : GetLicErrDesc = "The Software Licensing Service reported that the Genuine information contained in the license is not consistent."
    Case "0xC004E022" : GetLicErrDesc = "The Software Licensing Service reported that the secure store id value in license does not match with the current value."
    Case "0x8004E101" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store file version is invalid."
    Case "0x8004E102" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains an invalid descriptor table."
    Case "0x8004E103" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains a token with an invalid header/footer."
    Case "0x8004E104" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token has an invalid name."
    Case "0x8004E105" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token has an invalid extension."
    Case "0x8004E106" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store contains a duplicate token."
    Case "0x8004E107" : GetLicErrDesc = "The Software Licensing Service reported that a token in the Token Store has a size mismatch."
    Case "0x8004E108" : GetLicErrDesc = "The Software Licensing Service reported that a token in the Token Store contains an invalid hash."
    Case "0x8004E109" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store was unable to read a token."
    Case "0x8004E10A" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store was unable to write a token."
    Case "0x8004E10B" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store attempted an invalid file operation."
    Case "0x8004E10C" : GetLicErrDesc = "The Software Licensing Service reported that there is no active transaction."
    Case "0x8004E10D" : GetLicErrDesc = "The Software Licensing Service reported that the Token Store file header is invalid."
    Case "0x8004E10E" : GetLicErrDesc = "The Software Licensing Service reported that a Token Store token descriptor is invalid."
    Case "0xC004F001" : GetLicErrDesc = "The Software Licensing Service reported an internal error."
    Case "0xC004F002" : GetLicErrDesc = "The Software Licensing Service reported that rights consumption failed."
    Case "0xC004F003" : GetLicErrDesc = "The Software Licensing Service reported that the required license could not be found."
    Case "0xC004F004" : GetLicErrDesc = "The Software Licensing Service reported that the product key does not match the range defined in the license."
    Case "0xC004F005" : GetLicErrDesc = "The Software Licensing Service reported that the product key does not match the product key for the license."
    Case "0xC004F006" : GetLicErrDesc = "The Software Licensing Service reported that the signature file for the license is not available."
    Case "0xC004F007" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found."
    Case "0xC004F008" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found."
    Case "0xC004F009" : GetLicErrDesc = "The Software Licensing Service reported that the grace period expired."
    Case "0xC004F00A" : GetLicErrDesc = "The Software Licensing Service reported that the application ID does not match the application ID for the license."
    Case "0xC004F00B" : GetLicErrDesc = "The Software Licensing Service reported that the product identification data is not available."
    Case "0x4004F00C" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid grace period."
    Case "0x4004F00D" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid out of tolerance grace period."
    Case "0xC004F00E" : GetLicErrDesc = "The Software Licensing Service determined that the license could not be used by the current version of the security processor component."
    Case "0xC004F00F" : GetLicErrDesc = "The Software Licensing Service reported that the hardware ID binding is beyond the level of tolerance."
    Case "0xC004F010" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
    Case "0xC004F011" : GetLicErrDesc = "The Software Licensing Service reported that the license file is not installed."
    Case "0xC004F012" : GetLicErrDesc = "The Software Licensing Service reported that the call has failed because the value for the input key was not found."
    Case "0xC004F013" : GetLicErrDesc = "The Software Licensing Service determined that there is no permission to run the software."
    Case "0xC004F014" : GetLicErrDesc = "The Software Licensing Service reported that the product key is not available."
    Case "0xC004F015" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
    Case "0xC004F016" : GetLicErrDesc = "The Software Licensing Service determined that the request is not supported."
    Case "0xC004F017" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
    Case "0xC004F018" : GetLicErrDesc = "The Software Licensing Service reported that the license does not contain valid location data for the activation server."
    Case "0xC004F019" : GetLicErrDesc = "The Software Licensing Service determined that the requested event ID is invalid."
    Case "0xC004F01A" : GetLicErrDesc = "The Software Licensing Service determined that the requested event is not registered with the service."
    Case "0xC004F01B" : GetLicErrDesc = "The Software Licensing Service reported that the event ID is already registered."
    Case "0xC004F01C" : GetLicErrDesc = "The Software Licensing Service reported that the license is not installed."
    Case "0xC004F01D" : GetLicErrDesc = "The Software Licensing Service reported that the verification of the license failed."
    Case "0xC004F01E" : GetLicErrDesc = "The Software Licensing Service determined that the input data type does not match the data type in the license."
    Case "0xC004F01F" : GetLicErrDesc = "The Software Licensing Service determined that the license is invalid."
    Case "0xC004F020" : GetLicErrDesc = "The Software Licensing Service determined that the license package is invalid."
    Case "0xC004F021" : GetLicErrDesc = "The Software Licensing Service reported that the validity period of the license has expired."
    Case "0xC004F022" : GetLicErrDesc = "The Software Licensing Service reported that the license authorization failed."
    Case "0xC004F023" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
    Case "0xC004F024" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
    Case "0xC004F025" : GetLicErrDesc = "The Software Licensing Service reported that the action requires administrator privilege."
    Case "0xC004F026" : GetLicErrDesc = "The Software Licensing Service reported that the required data is not found."
    Case "0xC004F027" : GetLicErrDesc = "The Software Licensing Service reported that the license is tampered."
    Case "0xC004F028" : GetLicErrDesc = "The Software Licensing Service reported that the policy cache is invalid."
    Case "0xC004F029" : GetLicErrDesc = "The Software Licensing Service cannot be started in the current OS mode."
    Case "0xC004F02A" : GetLicErrDesc = "The Software Licensing Service reported that the license is invalid."
    Case "0xC004F02C" : GetLicErrDesc = "The Software Licensing Service reported that the format for the offline activation data is incorrect."
    Case "0xC004F02D" : GetLicErrDesc = "The Software Licensing Service determined that the version of the offline Confirmation ID (CID) is incorrect."
    Case "0xC004F02E" : GetLicErrDesc = "The Software Licensing Service determined that the version of the offline Confirmation ID (CID) is not supported."
    Case "0xC004F02F" : GetLicErrDesc = "The Software Licensing Service reported that the length of the offline Confirmation ID (CID) is incorrect."
    Case "0xC004F030" : GetLicErrDesc = "The Software Licensing Service determined that the Installation ID (IID) or the Confirmation ID (CID) could not been saved."
    Case "0xC004F031" : GetLicErrDesc = "The Installation ID (IID) and the Confirmation ID (CID) do not match. Please confirm the IID and reacquire a new CID if necessary."
    Case "0xC004F032" : GetLicErrDesc = "The Software Licensing Service determined that the binding data is invalid."
    Case "0xC004F033" : GetLicErrDesc = "The Software Licensing Service reported that the product key is not allowed to be installed. Please see the eventlog for details."
    Case "0xC004F034" : GetLicErrDesc = "The Software Licensing Service reported that the license could not be found or was invalid."
    Case "0xC004F035" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated with a Volume license product key. Volume-licensed systems require upgrading from a qualifying operating system. Please contact your system administrator or use a different type of key."
    Case "0xC004F038" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The count reported by your Key Management Service (KMS) is insufficient. Please contact your system administrator."
    Case "0xC004F039" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated.  The Key Management Service (KMS) is not enabled."
    Case "0x4004F040" : GetLicErrDesc = "The Software Licensing Service reported that the computer was activated but the owner should verify the Product Use Rights."
    Case "0xC004F041" : GetLicErrDesc = "The Software Licensing Service determined that the Key Management Service (KMS) is not activated. KMS needs to be activated. Please contact system administrator."
    Case "0xC004F042" : GetLicErrDesc = "The Software Licensing Service determined that the specified Key Management Service (KMS) cannot be used."
    Case "0xC004F047" : GetLicErrDesc = "The Software Licensing Service reported that the proxy policy has not been updated."
    Case "0xC004F04D" : GetLicErrDesc = "The Software Licensing Service determined that the Installation ID (IID) or the Confirmation ID (CID) is invalid."
    Case "0xC004F04F" : GetLicErrDesc = "The Software Licensing Service reported that license management information was not found in the licenses."
    Case "0xC004F050" : GetLicErrDesc = "The Software Licensing Service reported that the product key is invalid."
    Case "0xC004F051" : GetLicErrDesc = "The Software Licensing Service reported that the product key is blocked."
    Case "0xC004F052" : GetLicErrDesc = "The Software Licensing Service reported that the licenses contain duplicated properties."
    Case "0xC004F053" : GetLicErrDesc = "The Software Licensing Service determined that the license is invalid. The license contains an override policy that is not configured properly."
    Case "0xC004F054" : GetLicErrDesc = "The Software Licensing Service reported that license management information has duplicated data."
    Case "0xC004F055" : GetLicErrDesc = "The Software Licensing Service reported that the base SKU is not available."
    Case "0xC004F056" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated using the Key Management Service (KMS)."
    Case "0xC004F057" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
    Case "0xC004F058" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
    Case "0xC004F059" : GetLicErrDesc = "The Software Licensing Service reported that a license in the computer BIOS is invalid."
    Case "0xC004F060" : GetLicErrDesc = "The Software Licensing Service determined that the version of the license package is invalid."
    Case "0xC004F061" : GetLicErrDesc = "The Software Licensing Service determined that this specified product key can only be used for upgrading, not for clean installations."
    Case "0xC004F062" : GetLicErrDesc = "The Software Licensing Service reported that a required license could not be found."
    Case "0xC004F063" : GetLicErrDesc = "The Software Licensing Service reported that the computer BIOS is missing a required license."
    Case "0xC004F064" : GetLicErrDesc = "The Software Licensing Service reported that the non-genuine grace period expired."
    Case "0x4004F065" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid non-genuine grace period."
    Case "0xC004F066" : GetLicErrDesc = "The Software Licensing Service reported that the genuine information property can not be set before dependent property been set."
    Case "0xC004F067" : GetLicErrDesc = "The Software Licensing Service reported that the non-genuine grace period expired (type 2)."
    Case "0x4004F068" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid non-genuine grace period (type 2)."
    Case "0xC004F069" : GetLicErrDesc = "The Software Licensing Service reported that the product SKU is not found."
    Case "0xC004F06A" : GetLicErrDesc = "The Software Licensing Service reported that the requested operation is not allowed."
    Case "0xC004F06B" : GetLicErrDesc = "The Software Licensing Service determined that it is running in a virtual machine. The Key Management Service (KMS) is not supported in this mode."
    Case "0xC004F06C" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The Key Management Service (KMS) determined that the request timestamp is invalid."
    Case "0xC004F071" : GetLicErrDesc = "The Software Licensing Service reported that the plug-in manifest file is incorrect."
    Case "0xC004F072" : GetLicErrDesc = "The Software Licensing Service reported that the license policies for fast query could not be found."
    Case "0xC004F073" : GetLicErrDesc = "The Software Licensing Service reported that the license policies for fast query have not been loaded."
    Case "0xC004F074" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. No Key Management Service (KMS) could be contacted. Please see the Application Event Log for additional information."
    Case "0xC004F075" : GetLicErrDesc = "The Software Licensing Service reported that the operation cannot be completed because the service is stopping."
    Case "0xC004F076" : GetLicErrDesc = "The Software Licensing Service reported that the requested plug-in cannot be found."
    Case "0xC004F077" : GetLicErrDesc = "The Software Licensing Service determined incompatible version of authentication data."
    Case "0xC004F078" : GetLicErrDesc = "The Software Licensing Service reported that the key is mismatched."
    Case "0xC004F079" : GetLicErrDesc = "The Software Licensing Service reported that the authentication data is not set."
    Case "0xC004F07A" : GetLicErrDesc = "The Software Licensing Service reported that the verification could not be done."
    Case "0xC004F07B" : GetLicErrDesc = "The requested operation is unavailable while the Software Licensing Service is running."
    Case "0xC004F07C" : GetLicErrDesc = "The Software Licensing Service determined that the version of the computer BIOS is invalid."
    Case "0xC004F200" : GetLicErrDesc = "The Software Licensing Service reported that current state is not genuine."
    Case "0xC004F301" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The token-based activation challenge has expired."
    Case "0xC004F302" : GetLicErrDesc = "The Software Licensing Service reported that Silent Activation failed. The Software Licensing Service reported that there are no certificates found in the system that could activate the product without user interaction."
    Case "0xC004F303" : GetLicErrDesc = "The Software Licensing Service reported that the certificate chain could not be built or failed validation."
    Case "0xC004F304" : GetLicErrDesc = "The Software Licensing Service reported that required license could not be found."
    Case "0xC004F305" : GetLicErrDesc = "The Software Licensing Service reported that there are no certificates found in the system that could activate the product."
    Case "0xC004F306" : GetLicErrDesc = "The Software Licensing Service reported that this software edition does not support token-based activation."
    Case "0xC004F307" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation data is invalid."
    Case "0xC004F308" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation data is tampered."
    Case "0xC004F309" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. Activation challenge and response do not match."
    Case "0xC004F30A" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the conditions in the license."
    Case "0xC004F30B" : GetLicErrDesc = "The Software Licensing Service reported that the inserted smartcard could not be used to activate the product."
    Case "0xC004F30C" : GetLicErrDesc = "The Software Licensing Service reported that the token-based activation license content is invalid."
    Case "0xC004F30D" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The thumbprint is invalid."
    Case "0xC004F30E" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The thumbprint does not match any certificate."
    Case "0xC004F30F" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the criteria specified in the issuance license."
    Case "0xC004F310" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the trust point identifier (TPID) specified in the issuance license."
    Case "0xC004F311" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. A soft token cannot be used for activation."
    Case "0xC004F312" : GetLicErrDesc = "The Software Licensing Service reported that the computer could not be activated. The certificate cannot be used because its private key is exportable."
    Case "0xC004F313" : GetLicErrDesc = "The Software Licensing Service reported that the CNG encryption library could not be loaded.  The current certificate may not be available on this version of Windows."
    Case "0xC004FC03" : GetLicErrDesc = "A networking problem has occurred while activating your copy of Windows."
    Case "0x4004FC04" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the timebased validity period."
    Case "0x4004FC05" : GetLicErrDesc = "The Software Licensing Service reported that the application has a perpetual grace period."
    Case "0x4004FC06" : GetLicErrDesc = "The Software Licensing Service reported that the application is running within the valid extended grace period."
    Case "0xC004FC07" : GetLicErrDesc = "The Software Licensing Service reported that the validity period expired."
    Case "0xC004FE00" : GetLicErrDesc = "The Software Licensing Service reported that activation is required to recover from tampering of SL Service trusted store."
    Case "0xC004D101" : GetLicErrDesc = "The security processor reported an initialization error."
    Case "0x8004D102" : GetLicErrDesc = "The security processor reported that the machine time is inconsistent with the trusted time."
    Case "0xC004D103" : GetLicErrDesc = "The security processor reported that an error has occurred."
    Case "0xC004D104" : GetLicErrDesc = "The security processor reported that invalid data was used."
    Case "0xC004D105" : GetLicErrDesc = "The security processor reported that the value already exists."
    Case "0xC004D107" : GetLicErrDesc = "The security processor reported that an insufficient buffer was used."
    Case "0xC004D108" : GetLicErrDesc = "The security processor reported that invalid data was used."
    Case "0xC004D109" : GetLicErrDesc = "The security processor reported that an invalid call was made."
    Case "0xC004D10A" : GetLicErrDesc = "The security processor reported a version mismatch error."
    Case "0x8004D10B" : GetLicErrDesc = "The security processor cannot operate while a debugger is attached."
    Case "0xC004D301" : GetLicErrDesc = "The security processor reported that the trusted data store was tampered."
    Case "0xC004D302" : GetLicErrDesc = "The security processor reported that the trusted data store was rearmed."
    Case "0xC004D303" : GetLicErrDesc = "The security processor reported that the trusted store has been recreated."
    Case "0xC004D304" : GetLicErrDesc = "The security processor reported that entry key was not found in the trusted data store."
    Case "0xC004D305" : GetLicErrDesc = "The security processor reported that the entry key already exists in the trusted data store."
    Case "0xC004D306" : GetLicErrDesc = "The security processor reported that the entry key is too big to fit in the trusted data store."
    Case "0xC004D307" : GetLicErrDesc = "The security processor reported that the maximum allowed number of re-arms has been exceeded.  You must re-install the OS before trying to re-arm again."
    Case "0xC004D308" : GetLicErrDesc = "The security processor has reported that entry data size is too big to fit in the trusted data store."
    Case "0xC004D309" : GetLicErrDesc = "The security processor has reported that the machine has gone out of hardware tolerance."
    Case "0xC004D30A" : GetLicErrDesc = "The security processor has reported that the secure timer already exists."
    Case "0xC004D30B" : GetLicErrDesc = "The security processor has reported that the secure timer was not found."
    Case "0xC004D30C" : GetLicErrDesc = "The security processor has reported that the secure timer has expired."
    Case "0xC004D30D" : GetLicErrDesc = "The security processor has reported that the secure timer name is too long."
    Case "0xC004D30E" : GetLicErrDesc = "The security processor reported that the trusted data store is full."
    Case "0xC004D401" : GetLicErrDesc = "The security processor reported a system file mismatch error."
    Case "0xC004D402" : GetLicErrDesc = "The security processor reported a system file mismatch error."
    Case "0xC004D501" : GetLicErrDesc = "The security processor reported an error with the kernel data."
    Case "0xC004D502" : GetLicErrDesc = "Kernel Mode Cache is tampered and the restore attempt failed."
    Case "0x4004D601" : GetLicErrDesc = "Kernel Mode Cache was not changed."
    Case "0x4004D602" : GetLicErrDesc = "Reboot-requiring policies have changed."
    Case "0xC004D701" : GetLicErrDesc = "External decryption key was already set for specified feature."
    Case "0xC004D702" : GetLicErrDesc = "Error occured during proxy execution."
    Case "0x803FA065" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0x803FA066" : GetLicErrDesc = "The activation server determined there is a problem with the specified product key."
    Case "0x803FA067" : GetLicErrDesc = "The activation server determined the specified product key has been blocked."
    Case "0x803FA068" : GetLicErrDesc = "The activation server determined the specified product key is invalid."
    Case "0x803FA06C" : GetLicErrDesc = "The activation server determined the specified product key is unsupported."
    Case "0x803FA071" : GetLicErrDesc = "The activation server reported that the product key has exceeded its unlock limit."
    Case "0x803FA073" : GetLicErrDesc = "The activation server determined the license is invalid (INVALID_PRODUCT_DATA_ID)."
    Case "0x803FA074" : GetLicErrDesc = "The activation server determined the license is invalid (INVALID_PRODUCT_DATA)."
    Case "0x803FA076" : GetLicErrDesc = "The activation server determined the product key is not valid (INVALID_ACTCONFIG_ID)."
    Case "0x803FA077" : GetLicErrDesc = "The activation server determined the specified product key is invalid (INVALID_PRODUCT_KEY_LENGTH)."
    Case "0x803FA078" : GetLicErrDesc = "The activation server determined the specified product key is invalid (INVALID_PRODUCT_KEY_FORMAT)."
    Case "0x803FA07A" : GetLicErrDesc = "The activation server determined the specified product key is invalid (INVALID_ARGUMENT)."
    Case "0x803FA07D" : GetLicErrDesc = "The activation server experienced an error (INVALID_BINDING_URI)."
    Case "0x803FA07F" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key has exceeded its limit."
    Case "0x803FA080" : GetLicErrDesc = "The activation server reported that the Multiple Activation Key extension limit has been exceeded."
    Case "0x803FA083" : GetLicErrDesc = "The activation server reported that the specified product key cannot be used for online activation."
    Case "0x803FA08D" : GetLicErrDesc = "The activation server determined that the version of the offline Confirmation ID (CID) is incorrect."
    Case "0x803FA08E" : GetLicErrDesc = "The activation server reported that the format for the offline activation data is incorrect."
    Case "0x803FA08F" : GetLicErrDesc = "The activation server reported that the length of the offline Confirmation ID (CID) is incorrect."
    Case "0x803FA090" : GetLicErrDesc = "The Software Licensing Service determined that the Installation ID (IID) or the Confirmation ID (CID) is invalid."
    Case "0x803FA097" : GetLicErrDesc = "The activation server reported that time based activation attempted before start date."
    Case "0x803FA098" : GetLicErrDesc = "The activation server reported that time based activation attempted after end date."
    Case "0x803FA099" : GetLicErrDesc = "The activation server reported that new time based activation is not available."
    Case "0x803FA09A" : GetLicErrDesc = "The activation server reported that the time based product key is not configured for activation."
    Case "0x803FA0C8" : GetLicErrDesc = "The activation server reported that no business rules available to activate specified product key."
    Case "0x803FA0CB" : GetLicErrDesc = "The activation server determined the specified product key has been blocked for this geographic location."
    Case "0x803FA0D3" : GetLicErrDesc = "The activation server determined the license is invalid (INVALID_BINDING)."
    Case "0x803FA0D4" : GetLicErrDesc = "The activation server determined there is a problem with the specified product key (BINDING_NOT_CONFIGURED)."
    Case "0x803FA0D5" : GetLicErrDesc = "The activation server determined that the override limit is reached. (ROT)"
    Case "0x803FA0D6" : GetLicErrDesc = "The activation server determined that the override limit is reached."
    Case "0x803FA400" : GetLicErrDesc = "The activation server determined that the offer no longer exists."
    Case "0x803FABB8" : GetLicErrDesc = "Donor hardwareId does not own operating system entitlement."
    Case "0x803FABB9" : GetLicErrDesc = "User not eligible for reactivation. (Generic error code)"
    Case "0x803FABBA" : GetLicErrDesc = "User not eligible for reactivation due to no association."
    Case "0x803FABBB" : GetLicErrDesc = "User not eligible for reactivation because user is not an admin on device."
    Case "0x803FABBC" : GetLicErrDesc = "User not eligible for reactivation because user is is throttled."
    Case "0x803FABBD" : GetLicErrDesc = "User not eligible for reactivation because license is throttled."
    Case "0x803FABBE" : GetLicErrDesc = "Device is not eligible for reactivation because it is throttled."
    Case "0x803FABBF" : GetLicErrDesc = "User not eligible for because policy does not allow it."
    Case "0x803FABC0" : GetLicErrDesc = "Device is not eligible for reactivation because it is blocked."
    Case "0x803FABC1" : GetLicErrDesc = "User is not eligible for reactivation because the user is blocked."
    Case "0x803FABC2" : GetLicErrDesc = "License is not eligible for transfer because the license is blocked."
    Case "0x803FABC3" : GetLicErrDesc = "Device is not eligible for transfer because the device is blocked."
    Case Else : GetLicErrDesc = ""
    End Select
End Function 'GetLicErrDesc
'=======================================================================================================

'=======================================================================================================
'Module Productslist
'=======================================================================================================
Sub FindAllProducts
    Dim AllProducts, ProdX, key
    Dim arrKeys, Arr, arrTmpSids()
    Dim sSid
    Dim i
    On Error Resume Next
    
    'Cache ARP entries
    If RegEnumKey (HKLM, REG_ARP, arrKeys) Then
        For Each key in arrKeys
            If NOT dicArp.Exists(key) Then dicArp.Add key, REG_ARP & key
        Next 'key
    End If
    
    'Build an array of all applications registered to Windows Installer
    'Iterate products depending of available WI version
    Select Case iWiVersionMajor
    Case 3, 4, 5, 6
        sErrHnd = "_ErrorHandler3x" : Err.Clear
        
        Set AllProducts = oMsi.ProductsEx("", USERSID_EVERYONE, MSIINSTALLCONTEXT_ALL) : CheckError sActiveSub, sErrHnd 
        If CheckObject(AllProducts) Then WritePLArrayEx 3, arrAllProducts, AllProducts, Null, Null
    Case 2
        'Only available for backwards compatibility reasons
        'Will use direct registry reading. Reuse logic from FindProducts_ErrorHandler
        For i = 1 to 4
            'Note that 'i=3' is not a valid context which is simply ignored in FindRegProducts.
            FindRegProducts i
        Next 'i
    Case Else
        'Not Handled - Not Supported
    End Select
    
    GetC2RVersionActive

    FindRegProducts MSIINSTALLCONTEXT_C2RV2
    FindRegProducts MSIINSTALLCONTEXT_C2RV3

    InitMasterArray
    Set dicProducts = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(arrMaster)
        arrMaster(i, COL_ISOFFICEPRODUCT) = IsOfficeProduct(arrMaster(i, COL_PRODUCTCODE))
        If NOT dicProducts.Exists(arrMaster(i, COL_PRODUCTCODE)) Then dicProducts.Add arrMaster(i, COL_PRODUCTCODE), arrMaster(i, COL_USERSID)
    Next 'i
    'Build an array of all C2R virtualized applications 
    FindV1VirtualizedProducts
    FindV2VirtualizedProducts
    FindV3VirtualizedProducts
End Sub
'=======================================================================================================

Sub FindAllProducts_ErrorHandler3x
    
    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_PRODUCTSEXALL & DOT & " Error Details: " & _
    Err.Source & " " & Hex( Err ) & ": " & Err.Description
    
    On Error Resume Next
    
    FindProducts "", MSIINSTALLCONTEXT_MACHINE
    
    FindProducts USERSID_EVERYONE, MSIINSTALLCONTEXT_USERUNMANAGED
    
    FindProducts USERSID_EVERYONE, MSIINSTALLCONTEXT_USERMANAGED
End Sub
'=======================================================================================================

Sub FindProducts(sSid, iContext)
    Dim sActiveSub, sErrHnd 
    Dim Arr, Products
    sActiveSub = "FindProducts" : sErrHnd = ""
    On Error Resume Next

    sErrHnd = iContext : Err.Clear
    Set Products = oMsi.ProductsEx("", sSid, iContext) : CheckError sActiveSub, sErrHnd
    Select Case iContext
        Case MSIINSTALLCONTEXT_MACHINE :     If CheckObject(Products) Then WritePLArrayEx 3, arrMProducts, Products, iContext, sSid
        Case MSIINSTALLCONTEXT_USERUNMANAGED : If CheckObject(Products) Then WritePLArrayEx 3, arrUUProducts, Products, iContext, sSid
        Case MSIINSTALLCONTEXT_USERMANAGED : If CheckObject(Products) Then WritePLArrayEx 3, arrUMProducts, Products, iContext, sSid
        Case Else
    End Select
End Sub
'=======================================================================================================

Sub FindProducts_ErrorHandler (iContext)
    Dim sSid, Arr
    Dim n
    
    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, "ProductsEx for " & GetContextString(iContext) & " failed. Error Details: " & _
    Err.Source & " " & Hex( Err ) & ": " & Err.Description
    
    On Error Resume Next

    FindRegProducts iContext
End Sub
'=======================================================================================================

'General entry point for getting a list of products from the registry.
Sub FindRegProducts (iContext)
    Dim arrKeys, arrTmpKeys
    Dim sSid, sSubKeyName, sUMSids, sProd, sTmpProd, sVer
    Dim hDefKey
    Dim n, iProdCnt, iProdFind, iProdTotal
    On Error Resume Next

    Select Case iContext

    Case MSIINSTALLCONTEXT_MACHINE
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMProducts
    
    Case MSIINSTALLCONTEXT_C2RV2
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = REG_OFFICE & "15.0" & REG_C2RVIRT_HKLM & REG_GLOBALCONFIG & "S-1-5-18\Products\"
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMVProducts
    
    Case MSIINSTALLCONTEXT_C2RV3
        sSid = MACHINESID
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = REG_OFFICE & "ClickToRun\REGISTRY\MACHINE\" & REG_GLOBALCONFIG & "S-1-5-18\Products\"
        FindRegProductsEx hDefKey, sSubKeyName, sSid, iContext, arrMVProducts
    
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        sUMSids = ""
        If CheckArray (arrUMSids) Then sUMSids = Join(arrUMSids)
        If CheckArray (arrUUSids) Then
            iProdTotal = -1
            For iProdFind = 0 To 1
                For n = 0 To UBound(arrUUSids)
                    sSid = arrUUSids(n)
                    hDefKey = GetRegHive(iContext, sSid, False)
                    sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
                    If InStr(sUMSids, sSid) > 0 Then
                        'Current SID has installed managed per-user products
                        'Create a string list with products
                        sTmpProd = ""
                        If NOT sUMSids= "" Then
                            For iProdCnt = 0 To UBound(arrUMProducts)
                                If arrUMProducts(iProdCnt, COL_USERSID) =sSid Then sTmpProd = sTmpProd & arrUMProducts(iProdCnt, COL_PRODUCTCODE)
                            Next 'iProdCnt
                        End If
                        If RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys) Then
                            ReDim arrKeys(-1)
                            For iProdCnt = 0 To UBound(arrTmpKeys)
                                If Not InStr(sTmpProd, GetExpandedGuid(arrTmpKeys(iProdCnt))) > 0 Then
                                    ReDim Preserve arrKeys(UBound(arrKeys) + 1)
                                    arrKeys(UBound(arrKeys)) =arrTmpKeys(iProdCnt)
                                End If 'Not InStr(sTmpProd, arrTmpKeys(iProdCnt)) > 0
                            Next 'iProdCnt
                            If iProdFind = 0 Then
                                iProdTotal = iProdTotal + UBound(arrKeys) + 1
                            Else
                                WritePLArrayEx 0, arrUUProducts, arrKeys, iContext, sSid
                            End If
                        End If 'RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys)
                    Else
                        'No conflict with managed per user products
                        If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then
                            If iProdFind = 0 Then
                                iProdTotal = iProdTotal + UBound(arrKeys) + 1
                            Else
                                WritePLArrayEx 0, arrUUProducts, arrKeys, iContext, sSid
                            End If
                         End If 'RegEnumKey(hDefKey, sSubKeyName, arrKeys)
                    End If 'InStr(sUMSid, sSid) > 0
                Next 'n
                If iProdFind = 0 Then ReDim arrUUProducts(iProdFind, UBOUND_MASTER)
            Next 'iProdFind
        End If 'CheckArray

    Case MSIINSTALLCONTEXT_USERMANAGED
        iProdCnt = -1
        If CheckArray (arrUMSids) Then
            'Determine number of products
            For n = 0 To UBound(arrUMSids)
                sSid = arrUMSids (n)
                hDefKey = GetRegHive(iContext, sSid, False)
                sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
                If RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys) Then
                    iProdCnt = iProdCnt + UBound(arrTmpKeys) + 1
                End If
            Next 'n
            ReDim arrUMProducts(iProdCnt, UBOUND_MASTER)
                
            'Add the products to array
            For n = 0 To UBound(arrUMSids)
                sSid = arrUMSids (n)
                hDefKey = GetRegHive(iContext, sSid, False)
                sSubKeyName = GetRegConfigKey("", iContext, sSid, False)
                
                If RegEnumKey(hDefKey, sSubKeyName, arrTmpKeys) Then WritePLArrayEx 0, arrUMProducts, arrTmpKeys, iContext, sSid
            Next 'n
        End If 'CheckArray

    Case Else
    End Select
    
End Sub 'FindRegProducts
'=======================================================================================================

Sub FindRegProductsEx (hDefKey, sSubKeyName, sSid, iContext, Arr)
    Dim arrKeys
    On Error Resume Next

    If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then 
        WritePLArrayEx 0, Arr, arrKeys, iContext, sSid
    End If 'RegKeyExists
End Sub

'-------------------------------------------------------------------------------
'   FindV1VirtualizedProducts
'
'   Locate virtualized C2R_v1 products
'-------------------------------------------------------------------------------
Sub FindV1VirtualizedProducts
    Dim Key, sValue, VProd
    Dim arrKeys
    Dim dicVirtProd
    Dim iVCnt
    On Error Resume Next

    Set dicVirtProd = CreateObject("Scripting.Dictionary")

    For Each Key in dicArp.Keys
        If Len(Key) = 38 Then
            If IsOfficeProduct(Key) Then
                If RegReadValue(HKLM, REG_ARP & Key, "CVH", sValue, "REG_DWORD") Then
                    If sValue = "1" Then
                        If NOT dicVirtProd.Exists(Key) Then dicVirtProd.Add Key, Key
                    End If
                End If
            End If
        End If
    Next 'Key

    'Fill the virtual products array 
    If dicVirtProd.Count > 0 Then
        ReDim arrVirtProducts(dicVirtProd.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        For Each VProd in dicVirtProd.Keys
            'ProductCode
            arrVirtProducts(iVCnt, COL_PRODUCTCODE) = VProd
        
            'ProductName
            If RegReadValue(HKLM, REG_ARP & VProd, "DisplayName", sValue, "REG_SZ") Then
                arrVirtProducts(iVCnt , COL_PRODUCTNAME) = sValue
            End If

            'DisplayVersion
            If RegReadValue(HKLM, REG_ARP & VProd, "DisplayVersion", sValue, "REG_SZ") Then
                arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION) = sValue
            End If

            If IsOfficeProduct(VProd) Then
                'SP level
                If NOT fInitArrProdVer Then InitProdVerArrays
                arrVirtProducts(iVCnt, VIRTPROD_SPLEVEL) = OVersionToSpLevel (VProd, GetVersionMajor(VProd), arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION)) 
                'Architecture (bitness)
                arrVirtProducts(iVCnt, VIRTPROD_BITNESS) = "x86"
                If Mid(VProd, 21, 1) = "1" Then arrVirtProducts(iVCnt, VIRTPROD_BITNESS) = "x64"
            End If 'IsOfficeProduct

            iVCnt = iVCnt + 1
        Next 'VProd
    End If 'dicVirtProd > 0
End Sub 'FindV1VirtualizedProducts
'=======================================================================================================

'-------------------------------------------------------------------------------
'   FindV2VirtualizedProducts
'
'   Locate virtualized C2R_v2 products
'-------------------------------------------------------------------------------
Sub FindV2VirtualizedProducts
    Dim ArpItem, VProd, ConfigProd
    Dim component, culture, child, name, key, subKey, prod
    Dim sSubKeyName, sConfigName, sValue, sChild
    Dim sCurKey, sCurKeyL0, sCurKeyL1, sCurKeyL2, sCurKeyL3, sKey
    Dim sVersion, sVersionFallback, sFileName, sKeyComponents
    Dim sActive, sActiveConfiguration, sProd, sCult, sXNone, sSxsPath, sSxsExt
    Dim iVCnt
    Dim dicVirt2Prod, dicVirt2ConfigID
    Dim arrKeys, arrConfigProducts, arrVersion, arrCultures, arrChildPackages, arrSubKeys
    Dim arrNames, arrTypes, arrScenario, arrKeyComponents, arrComponentData
    Dim fUninstallString, fIsSP1, fC2RPolEnabled

	Dim REG_C2R				    : REG_C2R				   = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\"
   	Dim REG_C2RCONFIGURATION    : REG_C2RCONFIGURATION	   = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration\"
   	Dim REG_C2RPROPERTYBAG      : REG_C2RPROPERTYBAG       = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag\"
   	Dim REG_C2RSCENARIO         : REG_C2RSCENARIO          = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Scenario\"
   	Dim REG_C2RUPDATES          : REG_C2RUPDATES           = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Updates\"
	Dim REG_C2RPRODUCTIDS	    : REG_C2RPRODUCTIDS		   = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\"
	Dim REG_C2RUPDATEPOL	    : REG_C2RUPDATEPOL		   = "SOFTWARE\Policies\Microsoft\Office\15.0\Common\OfficeUpdate\"
    On Error Resume Next

    sXNone   = "x-none"
    sSxsPath = ""
    sSxsExt  = ""
    sActive  = "Active"
    
    If fv2SxS Then
	    REG_C2R				    = "SOFTWARE\Microsoft\Office\ClickToRun\"
   	    REG_C2RCONFIGURATION    = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration\"
   	    REG_C2RSCENARIO         = "SOFTWARE\Microsoft\Office\ClickToRun\Scenario\"
   	    REG_C2RUPDATES          = "SOFTWARE\Microsoft\Office\ClickToRun\Updates\15.0\"
	    REG_C2RPRODUCTIDS       = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\"
        sXNone   = "x-none.15"
        sSxsPath = "15.0\"
        sSxsExt  = ".15"
        If RegReadValue(HKLM, REG_C2RPRODUCTIDS, "ActiveConfiguration", sActiveConfiguration, REG_SZ) Then
            sActive = sActiveConfiguration
        End If
    End If

    Set dicVirt2Prod = CreateObject ("Scripting.Dictionary")
    Set dicVirt2ConfigID = CreateObject("Scripting.Dictionary")
    fIsSP1 = False
    fC2RPolEnabled = False
    ' extend ARP dic to contain C2R virtual references
    
    'Integration PackageGUID
    sPackageGuid = GetC2RIntegrationPackageGuid (dicC2RPropV2)

    'Config IDs
    AddC2RProductReleaseIds sSxsPath, dicVirt2ConfigID

    AddC2RConfigIds fv2SxS, sSxsPath, "15", REG_C2RPRODUCTIDS, dicVirt2ConfigID, dicC2RPropV2, dicVirt2Cultures

    ' enum ARP to identify configuration products
    AddC2RArpConfigIds "15", dicVirt2Prod, dicVirt2ConfigID

    'Fill the v2 virtual products array 
    If dicVirt2Prod.Count > 0 Then
        ReDim arrVirt2Products(dicVirt2Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        iVMVirtOverride = 15
        ReDim arrVirt2Products(dicVirt2Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        fIsSP1 = (CompareVersion(sVersionFallback, "15.0.4569.1506", True) > -1)
        fC2RPolEnabled = (CompareVersion(sVersionFallback, "15.0.4605.1003", True) > -1)
        
        'Global settings - applicable for all v2 products
        '------------------------------------------------
        AddC2RGlobalSettingsToArray fv2SxS, "15", REG_C2RSCENARIO, REG_C2RCONFIGURATION, REG_C2RPROPERTYBAG, sSxsPath, dicScenarioV2, dicVirt2Prod, dicC2RPropV2, arrVirt2Products
            
        'KeyComponentStates
        AddC2RKeyComponentStates dicKeyComponentsV2

        ' Loop for product specific details
        AddC2RProductDetailsToArray iVCnt, "15", REG_C2RPRODUCTIDS, dicVirt2ConfigID, dicVirt2Prod, dicC2RPropV2, arrVirt2Products
    End If 'dicVirtProd > 0
    iVMVirtOverride = 0

End Sub 'FindV2VirtualizedProducts

'-------------------------------------------------------------------------------
'   FindV3VirtualizedProducts
'
'   Locate virtualized C2R_v3 products
'-------------------------------------------------------------------------------
Sub FindV3VirtualizedProducts
    Dim ArpItem, VProd, culture, name, key, subKey, prod, component
    Dim sValue, sCurKey, sKey, sVersionFallback, sKeyComponents
    Dim sActiveConfiguration, sProd, sCult
    Dim iVCnt
    Dim dicVirt3Prod, dicVirt3ConfigID
    Dim arrKeys, arrConfigProducts, arrVersion, arrCultures
    Dim arrNames, arrTypes, arrSubKeys, arrKeyComponents, arrComponentData
    Dim fUninstallString
    
   	Const REG_C2RCONFIGURATION	    = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration\"
   	Const REG_C2RSCENARIO           = "SOFTWARE\Microsoft\Office\ClickToRun\Scenario\"
   	Const REG_C2RUPDATES            = "SOFTWARE\Microsoft\Office\ClickToRun\Updates\"
	Const REG_C2RPRODUCTIDS		    = "SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\"
	Const REG_C2R				    = "SOFTWARE\Microsoft\Office\ClickToRun\"
	Const REG_C2RUPDATEPOL		    = "SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate\"
    On Error Resume Next

    Set dicVirt3Prod = CreateObject ("Scripting.Dictionary")
    Set dicVirt3ConfigID = CreateObject("Scripting.Dictionary")
    ' extend ARP dic to contain C2R virt references
    AddC2RVirtualArp REG_C2R & "REGISTRY\MACHINE\" & REG_ARP
    
    'Integration PackageGUID
    sPackageGuid = GetC2RIntegrationPackageGuid (dicC2RPropV3)
    
    'ProductReleaseIds
    AddC2RProductReleaseIds "", dicVirt3ConfigID

    AddC2RConfigIds False, "","16", REG_C2RPRODUCTIDS, dicVirt3ConfigID, dicC2RPropV3, dicVirt3Cultures

    ' enum ARP to identify configuration products
    AddC2RArpConfigIds "16", dicVirt3Prod, dicVirt3ConfigID

    'Fill the v3 virtual products array 
    If dicVirt3Prod.Count > 0 Then
        ReDim arrVirt3Products(dicVirt3Prod.Count - 1, UBOUND_VIRTPROD)
        iVCnt = 0
        iVMVirtOverride = 16
        
        'Add global settings for v3 prods
        AddC2RGlobalSettingsToArray False, "16", REG_C2RSCENARIO, REG_C2RCONFIGURATION, "","", dicScenarioV3, dicVirt3Prod, dicC2RPropV3, arrVirt3Products

        'KeyComponentStates
        AddC2RKeyComponentStates dicKeyComponentsV3
            
        ' Loop for product specific details
        AddC2RProductDetailsToArray iVCnt, "16", REG_C2RPRODUCTIDS, dicVirt3ConfigID, dicVirt3Prod, dicC2RPropV3, arrVirt3Products
    End If 'dicVirtProd > 0
    iVMVirtOverride = 0

End Sub 'FindV3VirtualizedProducts


'-------------------------------------------------------------------------------
'   AddC2RProductDetailsToArray
'
'   Add product specific details
'-------------------------------------------------------------------------------
Sub AddC2RProductDetailsToArray (iVCnt, iVM, REG_C2RPRODUCTIDS, dicVirtConfigID, dicVirtProd, dicC2RProp, arrVirtProducts)
    Dim VProd, key, prod
    Dim sValue
    Dim fUninstallString
    On Error Resume Next

    For Each VProd in dicVirtProd.Keys

        'KeyName
        arrVirtProducts(iVCnt, VIRTPROD_KEYNAME) = VProd

        'ProductName
        If RegReadValue(HKLM, REG_ARP & VProd, "DisplayName", sValue, "REG_SZ") Then arrVirtProducts(iVCnt, COL_PRODUCTNAME) = sValue
        If NOT Len(sValue) > 0 Then arrVirtProducts(iVCnt, COL_PRODUCTNAME) = VProd

        'ConfigName
        arrVirtProducts (iVCnt, VIRTPROD_CONFIGNAME) = dicMapArpToConfigID.Item(VProd)
	        
        'DisplayVersion (BuildNumber)
        If RegReadValue(HKLM, REG_ARP & VProd, "DisplayVersion", sValue, "REG_SZ") Then 
            arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION) = sValue
        Else
            arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION) = dicC2RProp.item(STR_BUILDNUMBER)
        End If

        'C2R external Version reference
        If NOT fInitArrProdVer Then InitProdVerArrays
        arrVirtProducts (iVCnt, VIRTPROD_SPLEVEL) = GetC2RVersionFromBuild (arrVirtProducts(iVCnt, VIRTPROD_PRODUCTVERSION))

        iVCnt = iVCnt + 1
    Next 'VProd

End Sub 'AddC2RProductDetailsToArray


'-------------------------------------------------------------------------------
'   GetC2RVersionFromBuild
'
'   Get the version that maps to the C2R build
'-------------------------------------------------------------------------------
Function GetC2RVersionFromBuild (sBuild)
    Dim arrVersion
    Dim sReturn
    Dim iMinorHigh
    On Error Resume Next

    sReturn = ""

    If Len(sBuild) > 0 Then
        arrVersion = Split(sBuild, ".")

        Select Case arrVersion(0)
        Case "15"
        Case "16"
            Select Case arrVersion(2)
			Case "4229" : sReturn = "1509"
			Case "6001" : sReturn = "1509"
			Case "6228" : sReturn = "1510"
			Case "6326" : sReturn = "1511"
			Case "6366" : sReturn = "1511"
			Case "6430" : sReturn = "1512"
			Case "6470" : sReturn = "1512"
			Case "6528" : sReturn = "1601"
			Case "6568" : sReturn = "1601"
			Case "6701" : sReturn = "1602"
			Case "6741" : sReturn = "1602"
			Case "6729" : sReturn = "1603"
			Case "6769" : sReturn = "1603"
			Case "6828" : sReturn = "1604"
			Case "6868" : sReturn = "1604"
			Case "6925" : sReturn = "1605"
			Case "6965" : sReturn = "1605"
			Case "7030" : sReturn = "1606"
			Case "7070" : sReturn = "1606"
			Case "7127" : sReturn = "1607"
			Case "7167" : sReturn = "1607"
			Case "7301" : sReturn = "1608"
			Case "7341" : sReturn = "1608"
			Case "7329" : sReturn = "1609"
			Case "7369" : sReturn = "1609"
			Case "7426" : sReturn = "1610"
			Case "7466" : sReturn = "1610"
			Case "7531" : sReturn = "1611"
			Case "7571" : sReturn = "1611"
			Case "7628" : sReturn = "1612"
			Case "7668" : sReturn = "1612"
			Case "7726" : sReturn = "1701"
			Case "7766" : sReturn = "1701"
			Case "7830" : sReturn = "1702"
			Case "7870" : sReturn = "1702"
			Case "7927" : sReturn = "1703"
			Case "7967" : sReturn = "1703"
			Case "8027" : sReturn = "1704"
			Case "8067" : sReturn = "1704"
			Case "8201" : sReturn = "1705"
			Case "8229" : sReturn = "1706"
			Case "8326" : sReturn = "1707"
			Case "8431" : sReturn = "1708"
			Case "8528" : sReturn = "1709"
			Case "8625" : sReturn = "1710"
			Case "8730" : sReturn = "1711"
			Case "8827" : sReturn = "1712"
			Case "9001" : sReturn = "1801"
			Case "9029" : sReturn = "1802"
			Case "9126" : sReturn = "1803"
			Case "9226" : sReturn = "1804"
			Case "9330" : sReturn = "1805"
			Case "10228" : sReturn = "1806"
			Case "10325" : sReturn = "1807"
			Case "10336" : sReturn = "LTSB2018 - 201808"
			Case "10337" : sReturn = "LTSB2018 - 201809"
			Case "10338" : sReturn = "LTSB2018 - 201810"
			Case "10339" : sReturn = "LTSB2018 - 201811"
			Case "10340" : sReturn = "LTSB2018 - 201812"
			Case "10341" : sReturn = "LTSB2018 - 201901"
			Case "10342" : sReturn = "LTSB2018 - 201902"
			Case "10343" : sReturn = "LTSB2018 - 201903"
			Case "10344" : sReturn = "LTSB2018 - 201904"
			Case "10345" : sReturn = "LTSB2018 - 201905"
			Case "10346" : sReturn = "LTSB2018 - 201906"
			Case "10347" : sReturn = "LTSB2018 - 201907"
			Case "10348" : sReturn = "LTSB2018 - 201908"
			Case "10349" : sReturn = "LTSB2018 - 201909"
			Case "10350" : sReturn = "LTSB2018 - 201910"
			Case "10351" : sReturn = "LTSB2018 - 201911"
			Case "10352" : sReturn = "LTSB2018 - 201912"
			Case "10353" : sReturn = "LTSB2018 - 202001"
			Case "10730" : sReturn = "1808"
			Case "10827" : sReturn = "1809"
			Case "11001" : sReturn = "1810"
			Case "11029" : sReturn = "1811"
			Case "11126" : sReturn = "1812"
			Case "11231" : sReturn = "1901"
			Case "11328" : sReturn = "1902"
			Case "11425" : sReturn = "1903"
			Case "11601" : sReturn = "1904"
			Case "11629" : sReturn = "1905"
			Case "11727" : sReturn = "1906"
			Case "11901" : sReturn = "1907"
			Case "11929" : sReturn = "1908"
			Case "12026" : sReturn = "1909"
			Case "12130" : sReturn = "1910"
			Case "12228" : sReturn = "1911"
			Case "12325" : sReturn = "1912"
			Case "12430" : sReturn = "2001"
			Case "12527" : sReturn = "2002"
            Case Else
                iMinorHigh = CInt(arrVersion(2))
                If iMinorHigh > 8827 AND iMinorHigh < 9001 Then sReturn = "1801 (DevMain)"
                If iMinorHigh > 9001 AND iMinorHigh < 9029 Then sReturn = "1802 (DevMain)"
                If iMinorHigh > 9029 AND iMinorHigh < 9126 Then sReturn = "1803 (DevMain)"
                If iMinorHigh > 9126 AND iMinorHigh < 9226 Then sReturn = "1804 (DevMain)"
                If iMinorHigh > 9226 AND iMinorHigh < 9330 Then sReturn = "1805 (DevMain)"
                If iMinorHigh > 9330 AND iMinorHigh < 10228 Then sReturn = "1806 (DevMain)"
                If iMinorHigh > 10228 AND iMinorHigh < 10325 Then sReturn = "1807 (DevMain)"
                If iMinorHigh > 10325 AND iMinorHigh < 10336 Then sReturn = "1808 (DevMain)"
                If iMinorHigh > 10336 AND iMinorHigh < 10730 Then sReturn = "LTSB 2018"
                If iMinorHigh > 10730 AND iMinorHigh < 10827 Then sReturn = "1809 (DevMain)"
                If iMinorHigh > 10827 AND iMinorHigh < 11001 Then sReturn = "1810 (DevMain)"
                If iMinorHigh > 11001 AND iMinorHigh < 11029 Then sReturn = "1811 (DevMain)"
                If iMinorHigh > 11029 AND iMinorHigh < 11126 Then sReturn = "1812 (DevMain)"
                If iMinorHigh > 11126 AND iMinorHigh < 11231 Then sReturn = "1901 (DevMain)"
                If iMinorHigh > 11231 AND iMinorHigh < 11328 Then sReturn = "1902 (DevMain)"
                If iMinorHigh > 11328 AND iMinorHigh < 11425 Then sReturn = "1903 (DevMain)"
                If iMinorHigh > 11425 AND iMinorHigh < 11601 Then sReturn = "1904 (DevMain)"
                If iMinorHigh > 11601 AND iMinorHigh < 11629 Then sReturn = "1905 (DevMain)"
                If iMinorHigh > 11629 AND iMinorHigh < 11727 Then sReturn = "1906 (DevMain)"
                If iMinorHigh > 11727 AND iMinorHigh < 11901 Then sReturn = "1907 (DevMain)"
                If iMinorHigh > 11901 AND iMinorHigh < 11929 Then sReturn = "1908 (DevMain)"
                If iMinorHigh > 11929 AND iMinorHigh < 12026 Then sReturn = "1909 (DevMain)"
                If iMinorHigh > 12026 AND iMinorHigh < 12130 Then sReturn = "1910 (DevMain)"
                If iMinorHigh > 12130 AND iMinorHigh < 12228 Then sReturn = "1911 (DevMain)"
                If iMinorHigh > 12228 AND iMinorHigh < 12325 Then sReturn = "1912 (DevMain)"
                If iMinorHigh > 12325 AND iMinorHigh < 12430 Then sReturn = "2001 (DevMain)"
                If iMinorHigh > 12430 AND iMinorHigh < 12501 Then sReturn = "2002 (DevMain)"
            End Select
        End Select
    End If

    GetC2RVersionFromBuild = sReturn
End Function 'GetC2RVersionFromBuild

'-------------------------------------------------------------------------------
'   AddC2RVirtualArp
'
'   Extend the ARP dictionary with the virtual registry keys
'-------------------------------------------------------------------------------
Sub AddC2RVirtualArp (sKey)
    Dim sValue
    Dim key
    Dim arrKeys
	Const REG_C2R				    = "SOFTWARE\Microsoft\Office\ClickToRun\"
    On Error Resume Next

    If RegEnumKey (HKLM, sKey, arrKeys) Then
        For Each key in arrKeys
            If NOT dicArp.Exists(key) Then dicArp.Add key, sKey & key
        Next 'key
    End If

End Sub 'AddC2RVirtualArp

'-------------------------------------------------------------------------------
'   GetC2RIntegrationPackageGuid
'
'   Obtain active integration PackageGuid vor C2R
'-------------------------------------------------------------------------------
Function GetC2RIntegrationPackageGuid (dicC2RProp)
    Dim sValue
    Dim arrKeys
	Const REG_C2R				    = "SOFTWARE\Microsoft\Office\ClickToRun\"
    On Error Resume Next

    sPackageGuid = ""
    If RegReadValue(HKLM, REG_C2R, "PackageGUID", sValue, REG_SZ) Then
    	sPackageGuid = "{" & sValue & "}"
        If NOT dicC2RProp.Exists(STR_REGPACKAGEGUID) Then dicC2RProp.Add STR_REGPACKAGEGUID, sValue
        If NOT dicC2RProp.Exists(STR_PACKAGEGUID) Then dicC2RProp.Add STR_PACKAGEGUID, sPackageGuid
    Else
        If RegEnumKey(HKLM, REG_C2R & "appvMachineRegistryStore\Integration\Packages\", arrKeys) Then
    	    sPackageGuid = sValue
    	    sValue = Replace(sValue, "{","")
    	    sValue = Replace(sValue, "}","")
            If NOT dicC2RProp.Exists(STR_REGPACKAGEGUID) Then dicC2RProp.Add STR_REGPACKAGEGUID, sValue
            If NOT dicC2RProp.Exists(STR_PACKAGEGUID) Then dicC2RProp.Add STR_PACKAGEGUID, sPackageGuid
        End If
    End If

    GetC2RIntegrationPackageGuid = sPackageGuid
End Function 'GetC2RIntegrationPackageGuid

'-------------------------------------------------------------------------------
'   AddC2RProductReleaseIds
'
'   Adds C2R config id's from ProductReleaseIds to ConfigId dictionary
'-------------------------------------------------------------------------------
Sub AddC2RProductReleaseIds (sSxsPath, dicC2RVirtConfigId)
    Dim sValue
    Dim prod
	Const REG_C2RCONFIGURATION    = "SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    On Error Resume Next

    If RegReadValue(HKLM, REG_C2RCONFIGURATION & sSxsPath, "ProductReleaseIds", sValue, REG_SZ) Then
        For Each prod in Split(sValue, ",")
            If NOT dicC2RVirtConfigId.Exists(prod) Then
                dicC2RVirtConfigId.Add prod, prod
            End If
        Next 'prod
    End If

End Sub 'AddC2RProductReleaseIds

'-------------------------------------------------------------------------------
'   AddC2RConfigIds
'
'   Adds C2R config id's and cultures 
'   from ProductReleaseId subkeys to ConfigId dictionary
'-------------------------------------------------------------------------------
Sub AddC2RConfigIds (fv2SxS, sSxsPath, sVersionMajor, sReg_C2RProductIds, dicC2RVirtConfigId, dicC2RProp, dicC2RCultures)
    Dim sProd, sActiveConfiguration, sXNoneCultureVer, sVersionFallback, sCult
    Dim prod, culture
    Dim arrConfigProducts, arrCultures
    Dim fv2NoSxS
    On Error Resume Next

    fv2NoSxS = (Not fv2SxS AND (sVersionMajor = "15"))

    If RegReadValue(HKLM, sReg_C2RProductIds, "ActiveConfiguration", sActiveConfiguration, REG_SZ) Then
        If RegEnumKey(HKLM, sReg_C2RProductIds & sActiveConfiguration, arrConfigProducts) Then
    	    For Each prod In arrConfigProducts
    		    sProd = prod
			    If InStr(sProd, ".1") > 0 Then sProd = Left(sProd, InStr(sProd, ".1") - 1)
    		    Select Case LCase(sProd)
    		    Case "culture","stream"
    		    Case Else
	                'add to ConfigID collection
	                If NOT dicC2RVirtConfigId.Exists(sProd) Then
	                    dicC2RVirtConfigId.Add sProd, prod
	                End If
    		    End Select
    	    Next 'prod
        End If 'arrConfigProducts

        'Shared ProductVersion (BuildNumber) Fallback
        sXNoneCultureVer = "\culture\x-none." & sVersionMajor & "\"
        If fv2NoSxS  Then sXNoneCultureVer = "\culture\"
        If RegReadStringValue(HKLM, sReg_C2RProductIds & sActiveConfiguration & sXNoneCultureVer, "Version", sVersionFallback) Then
            AddToDic dicC2RProp, STR_BUILDNUMBER, sVersionFallback
        End If
    	
    	'Cultures
        If RegEnumKey(HKLM, sReg_C2RProductIds & sActiveConfiguration & "\culture", arrCultures) Then
    		For Each culture in arrCultures
    			sCult = culture
				If InStr(sCult, ".1") > 0 Then sCult = Left(sCult, InStr(sCult, ".1") - 1)
    			Select Case LCase(sCult)
    			Case "x-none"
    			Case Else
	                'add to ConfigID collection
                    ' LP's and LIP's do not have a dedicated identifier
	                If NOT dicC2RCultures.Exists(sCult) Then
	                	dicC2RCultures.Add sCult, culture
	                End If
    			End Select
    		Next 'culture
        End If 'cultures
    End If 'ActiveConfiguration

End Sub 'AddC2RConfigIds

'-------------------------------------------------------------------------------
'   AddC2RArpConfigIds
'
'   Adds C2R config id's from ARP 
'-------------------------------------------------------------------------------
Sub AddC2RArpConfigIds (sVM, dicVirtProd, dicVirtConfigID)
    Dim ArpItem, key, prod
    Dim sCurKey, sValue, sDotVM, sVM_, sMSOfficeVer
    Dim fUninstallString
    On Error Resume Next

    sDotVM = "." & sVM
    sVM_ = "." & sVM & "_"
    sMSOfficeVer = "microsoft office " & sVM
    ' enum ARP to identify configuration products
    For Each ArpItem in dicArp.Keys
        ' filter on C2Rvx products
        sCurKey = REG_ARP & ArpItem & "\"
        fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sValue, "REG_SZ")
        If InStr(LCase(sValue), sMSOfficeVer) > 0 OR _
           (InStr(LCase(sValue), "productstoremove=") > 0 AND _
           InStr(sValue, sVM_) > 0) Then
        	For Each key In dicVirtConfigID.Keys
        		If InStr(sValue, key) > 0 Then
		            AddToDic dicVirtProd, ArpItem, sCurKey
                    AddToDic dicMapArpToConfigID, ArpItem, key
        		End If
        	Next
            If InStr(sValue, "productsdata ") > 0 Then prod = Mid(sValue, InStr(sValue, "productsdata ") + 13)
            prod = Trim(prod)
            If InStr(sValue, "productstoremove=") > 0 Then
        	    prod = Mid(sValue, InStr(sValue, "productstoremove="))
	            prod = Replace(prod, "productstoremove=","")
	        End If
            If InStr(prod, "_") > 0 Then  prod = Left(prod, InStr(prod, "_") - 1)
	        If InStr(prod, sDotVM) > 0 Then
	            prod = Left(prod, InStr(prod, sDotVM) - 1)
	            AddToDic dicVirtProd, ArpItem, sCurKey
                AddToDic dicMapArpToConfigID, ArpItem, key
	        End If
        End If
    Next 'ArpItem
End Sub 'AddC2RArpConfigIds

'-------------------------------------------------------------------------------
'   AddC2RGlobalSettingsToArray
'
'   Global / non-product specific settings for C2R products
'-------------------------------------------------------------------------------
Sub AddC2RGlobalSettingsToArray (fv2SxS, sVM, sRegScenario, sRegConfiguration, sRegPropertyBag, sSxsPath, dicScenario, dicVirtProd, dicC2RProp, arrVirtProducts)
    On Error Resume Next

    'Build and Version numbers 
    AddC2RVersionDetails fv2Sxs, sVM, sRegConfiguration, dicC2RProp
    
    'Scenario key state(s)
    AddC2RScenarioStates sRegScenario, dicScenario

    'Architecture (bitness)
    AddC2RBitness sRegConfiguration, sRegPropertyBag, dicC2RProp

    'SCA
    AddC2RSCA sVM, sRegConfiguration, dicC2RProp

    'CDNBaseUrl
    AddC2RCDNBaseUrl fv2SxS, sVM, sRegScenario, sRegConfiguration, sRegPropertyBag, sSxsPath, dicC2RProp

    'Updates Configuration
    AddC2RUpdateSettings sVM, sRegConfiguration, dicC2RProp

    'OfficeMgmtCom
    If sVM = "16" Then AddC2ROfficeMgmtCom sVM, sRegConfiguration, dicC2RProp 

End Sub 'AddC2RGlobalSettingsToArray

'-------------------------------------------------------------------------------
'   AddC2RVersionDetails
'
'   C2R BuildNumber and external Version reference
'-------------------------------------------------------------------------------
Sub AddC2RVersionDetails (fv2SxS, sVM, sRegConfiguration, dicC2RProp)
    Dim sVersionToReport
    Dim key
    On Error Resume Next

    'Global dicActiveC2RVersions
    'ProductVersionToReport
    For Each key in dicActiveC2RVersions.Keys
        If Left(key, 2) = sVM Then AddToDic dicC2RProp, STR_BUILDNUMBER, key
        'Translate BuildNumber to Version reference
        AddToDic dicC2RProp, STR_VERSION, GetC2RVersionFromBuild(key)
    Next
    'Shared ProductVersionToReport (BuildNumber)
    If RegReadValue (HKLM, sRegConfiguration, "VersionToReport", sVersionToReport, "REG_SZ") Then
        AddToDic dicC2RProp, STR_VERSIONTOREPORT, sVersionToReport
    End If
End Sub 'AddC2RVersionDetails

'-------------------------------------------------------------------------------
'   AddC2RScenarioStates
'
'   Scenario key state(s)
'-------------------------------------------------------------------------------
Sub AddC2RScenarioStates (sRegScenario, dicScenario)
    Dim arrNames, arrTypes, arrKeys, arrSubKeys
    Dim name, key, subKey, sValue
    On Error Resume Next

    If RegEnumKey(HKLM, sRegScenario, arrKeys) Then
        For Each key in arrKeys
            If RegEnumKey (HKLM, sRegScenario & key, arrSubKeys) Then
                For Each subKey in arrSubKeys
                    If RegEnumValues(HKLM, sRegScenario & key & "\" & subKey, arrNames, arrTypes) Then
                        For Each name in arrNames
                            RegReadValue HKLM, sRegScenario & key & "\" & subKey, name, sValue, "REG_SZ"
                            If InStr (name, ":") > 0 Then name = Left (name, InStr(name , ":") - 1)
                            AddToDic dicScenario, key & "\" & name, sValue
                        Next 'name
                    End If
                Next 'subKey
            End If
        Next 'name
    Else
        If RegEnumValues(HKLM, sRegScenario, arrNames, arrTypes) Then
            For Each name in arrNames
                RegReadValue HKLM, sRegScenario, name, sValue, "REG_DWORD"
                AddToDic dicScenario, name, sValue
            Next 'name
        End If
    End If

End Sub 'AddC2RScenarioStates

'-------------------------------------------------------------------------------
'   AddC2RBitness
'
'   C2R Architecture / Bitness
'-------------------------------------------------------------------------------
Sub AddC2RBitness (sRegConfiguration, sRegPropertyBag, dicC2RProp)
    Dim sValue
    On Error Resume Next

    If RegReadValue (HKLM, sRegConfiguration, "Platform", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_PLATFORM, sValue
    ElseIf RegReadValue (HKLM, sRegPropertyBag, "platform", sValue, "REG_SZ") Then 
        AddToDic dicC2RProp, STR_PLATFORM, sValue
    Else
        AddToDic dicC2RProp, STR_PLATFORM, "Error"
    End If

End Sub 'AddC2RBitness

'-------------------------------------------------------------------------------
'   AddC2RSCA
'
'   Check if SCA is enabled
'-------------------------------------------------------------------------------
Sub AddC2RSCA (sVM, sRegConfiguration, dicC2RProp)
    Const REG_LMPOLICIESOFFICELICENSING = "Software\Policies\Microsoft\Office\16.0\Common\Licensing"
    Dim sValue
    On Error Resume Next

    ' as obtained from C2R Configuration key
    If RegReadValue (HKLM, sRegConfiguration, "SharedComputerLicensing", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_SCA, sValue
    ElseIf (sVM = "16" AND RegKeyExists(HKLM, REG_LMPOLICIESOFFICELICENSING)) Then
        If RegReadValue (HKLM, REG_LMPOLICIESOFFICELICENSING, "SharedComputerLicensing", sValue, "REG_DWORD") Then
            AddOrUpdateDic dicC2RProp, STR_SCA, sValue
        End If
        If RegReadValue (HKLM, REG_LMPOLICIESOFFICELICENSING, "SCLCacheOverride", sValue, "REG_SZ") Then
            AddOrUpdateDic dicC2RProp, STR_POLSCACACHEOVERRIDE, sValue
        End If
        If RegReadValue (HKLM, STR_POLSCACACHEOVERRIDEDIR, "SCLCacheOverrideDirectory", sValue, "REG_SZ") Then
            AddOrUpdateDic dicC2RProp, STR_POLSCACACHEOVERRIDEDIR, sValue
        End If

    End If
    If Not dicC2RProp.Exists(STR_SCA) Then dicC2RProp.Add STR_SCA, STR_NOTCONFIGURED

End Sub 'AddC2RSCA

'-------------------------------------------------------------------------------
'   AddC2RCDNBaseUrl
'
'   Add value for the registered CDNBaseUrl
'-------------------------------------------------------------------------------
Sub AddC2RCDNBaseUrl (fv2SxS, sVM, sRegScenario, sRegConfiguration, sRegPropertyBag, sSxsPath, dicC2RProp)
    Dim sValue, sISC
    On Error Resume Next

    sISC = hAtS(Array("43","44","4E","20","4E","61","6D","65"))

    If RegReadValue(HKLM, sRegConfiguration & sSxsPath, "CDNBaseUrl", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_CDNBASEURL, sValue
        AddBrDet sValue, dicC2RProp
        'Last known InstallSource used
        If RegReadValue (HKLM, sRegScenario & "Install\","BaseUrl", sValue, "REG_SZ") Then
            AddToDic dicC2RProp, STR_LASTUSEDBASEURL, sValue
        End If 
    ElseIf RegReadValue(HKLM, sRegScenario & "INSTALL\" & sSxsPath, "BaseUrl", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_CDNBASEURL, sValue
        AddBrDet sValue, dicC2RProp
        'Last known InstallSource used
        AddToDic dicC2RProp, STR_LASTUSEDBASEURL, sValue
    ElseIf RegReadValue (HKLM, sRegPropertyBag, "wwwbaseurl", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_CDNBASEURL, sValue
        AddBrDet sValue, dicC2RProp
    Else
        If fv2SxS Then
            AddToDic dicC2RProp, STR_CDNBASEURL, "Default"
        Else
            AddToDic dicC2RProp, STR_CDNBASEURL, "Error"
        End If
    End If

End Sub 'AddC2RCDNBaseUrl

'-------------------------------------------------------------------------------
'   AddC2RUpdateSettings
'
'   Collect details for update settings
'-------------------------------------------------------------------------------
Sub AddC2RUpdateSettings (sVM, sRegConfiguration, dicC2RProp)
    Dim sValue, sRegC2RUpdatePol, sRegC2RUpdates
    On Error Resume Next

    sRegC2RUpdatePol = "SOFTWARE\Policies\Microsoft\Office\" & sVM & ".0\Common\OfficeUpdate\"
    sRegC2RUpdates = Replace(sRegConfiguration, "Configuration","Updates")

    'UpdatesEnabled - Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "enableautomaticupdates", sValue, "REG_DWORD") Then
        AddToDic dicC2RProp, STR_POLUPDATESENABLED, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATESENABLED, STR_NOTCONFIGURED
    End If

    'UpdateChannel - Setting
    If RegReadValue (HKLM, sRegConfiguration, "UpdateChannel", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_UPDATECHANNEL, sValue
    ElseIf RegReadValue (HKLM, sRegConfiguration, "UpdateBranch", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_UPDATECHANNEL, sValue
    Else
        AddToDic dicC2RProp, STR_UPDATECHANNEL, STR_NOTCONFIGURED
    End If

    'UpdateChannel - Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "updatebranch", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLUPDATECHANNEL, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATECHANNEL, STR_NOTCONFIGURED
    End If

    'UpdateUrl / Path - Setting
    If RegReadValue (HKLM, sRegConfiguration, "UpdateUrl", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_UPDATELOCATION, sValue
    Else
        AddToDic dicC2RProp, STR_UPDATELOCATION, STR_NOTCONFIGURED
    End If

    'UpdateUrl / Path Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "updatepath", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLUPDATELOCATION, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATELOCATION, STR_NOTCONFIGURED
    End If

    'Default to CDNBaseUrl - Setting
    sValue = dicC2RProp.Item (STR_CDNBASEURL)
    If NOT dicC2RProp.Item (STR_UPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RProp.Item (STR_UPDATELOCATION)
    If NOT dicC2RProp.Item (STR_POLUPDATECHANNEL) = STR_NOTCONFIGURED Then sValue = dicC2RProp.Item (STR_POLUPDATECHANNEL)
    If NOT dicC2RProp.Item (STR_POLUPDATELOCATION) = STR_NOTCONFIGURED Then sValue = dicC2RProp.Item (STR_POLUPDATELOCATION)
    If UCASE (dicC2RProp.Item (STR_UPDATESENABLED)) = "FALSE" Then sValue = "-"
    If UCASE (dicC2RProp.Item (STR_POLUPDATESENABLED)) = "FALSE" Then sValue = "-"
    If NOT GetBr (sValue) = "Custom" Then sValue = sValue & " (" & GetBr (sValue) & ")"
    dicC2RProp.Add STR_USEDUPDATELOCATION, sValue

    'UpdateVersion - Setting
    If RegReadValue (HKLM, sRegConfiguration, "UpdateToVersion", sValue, "REG_SZ") Then
        If sValue = "" Then sValue = STR_NOTCONFIGURED
        AddToDic dicC2RProp, STR_UPDATETOVERSION, sValue
    ElseIf RegReadValue (HKLM, sRegC2RUpdates, "UpdateToVersion", sValue, "REG_SZ") Then
        If sValue = "" Then sValue = STR_NOTCONFIGURED
        AddToDic dicC2RProp, STR_UPDATETOVERSION, sValue
    Else
        AddToDic dicC2RProp, STR_UPDATETOVERSION, STR_NOTCONFIGURED
    End If
        
    'UpdateVersion - Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "updatetargetversion", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLUPDATETOVERSION, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATETOVERSION, STR_NOTCONFIGURED
    End If

    'UpdateDeadline Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "updatedeadline", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLUPDATEDEADLINE, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATEDEADLINE, STR_NOTCONFIGURED
    End If

    'UpdateHideEnableDisableUpdates Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "hideenabledisableupdates ", sValue, "REG_DWORD") Then
        AddToDic dicC2RProp, STR_POLHIDEUPDATECFGOPT, sValue
    Else
        AddToDic dicC2RProp, STR_POLHIDEUPDATECFGOPT, STR_NOTCONFIGURED
    End If
            
    'UpdateNotifications Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "hideupdatenotifications ", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLUPDATENOTIFICATIONS, sValue
    Else
        AddToDic dicC2RProp, STR_POLUPDATENOTIFICATIONS, STR_NOTCONFIGURED
    End If

    'UpdateThrottle
    If RegReadValue (HKLM, sRegC2RUpdates, "UpdatesThrottleValue", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_UPDATETHROTTLE, sValue
    Else
        AddToDic dicC2RProp, STR_UPDATETHROTTLE, ""
    End If


End Sub 'AddC2RUpdateSettings

'-------------------------------------------------------------------------------
'   AddC2ROfficeMgmtCom
'
'   Check if OfficeMgmtCom (SCCM integration) is enabled
'-------------------------------------------------------------------------------
Sub AddC2ROfficeMgmtCom (sVM, sRegConfiguration, dicC2RProp)
    Dim sValue, sRegC2RUpdatePol
    On Error Resume Next

    sRegC2RUpdatePol = "SOFTWARE\Policies\Microsoft\Office\" & sVM & ".0\Common\OfficeUpdate\"

    'OfficeMgmtCom - Setting
    If RegReadValue (HKLM, sRegConfiguration, "OfficeMgmtCom", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_OFFICEMGMTCOM, sValue
    Else
        AddToDic dicC2RProp, STR_OFFICEMGMTCOM, STR_NOTCONFIGURED
    End If

    'OfficeMgmtCom - Policy
    If RegReadValue (HKLM, sRegC2RUpdatePol, "OfficeMgmtCom", sValue, "REG_SZ") Then
        AddToDic dicC2RProp, STR_POLOFFICEMGMTCOM, sValue
    Else
        AddToDic dicC2RProp, STR_POLOFFICEMGMTCOM, STR_NOTCONFIGURED
    End If

End Sub 'AddC2ROfficeMgmtCom

'-------------------------------------------------------------------------------
'   AddC2RKeyComponentStates
'
'   Adds KeyComponent states for core application components
'-------------------------------------------------------------------------------
Sub AddC2RKeyComponentStates (dicKeyComponents)
    Dim component
    Dim sKeyComponents
    Dim arrKeyComponents, arrComponentData
    On Error Resume Next

    If NOT sPackageGuid = "" Then
        sKeyComponents = GetKeyComponentStates(sPackageGuid, False)
        arrKeyComponents = Split(sKeyComponents, ";")
        For Each component in arrKeyComponents
            arrComponentData = Split(component, ",")
            If CheckArray(arrComponentData) Then
                AddToDic dicKeyComponents, arrComponentData(0), component
            End If
        Next 'component
    End If

End Sub 'AddC2RKeyComponentStates

'-------------------------------------------------------------------------------
'   MapAppsToConfigProduct
'
'   Obtain the mapping of which apps are contained in a SKU based on the
'   products ConfigID
'   Pass in a dictionary object which and return the filled dic 
'-------------------------------------------------------------------------------
Sub MapAppsToConfigProduct(dicConfigIdApps, sConfigId, iVM)
    Dim sMondo

    dicConfigIdApps.RemoveAll
    'sMondo = "MondoVolume"
    sMondo = hAtS(Array("4D","6F","6E","64","6F","56","6F","6C","75","6D","65"))
    Select Case sConfigId
    Case sMondo
        dicConfigIdApps.Add "Access",""
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "Groove",""
        If iVM = 15 Then dicConfigIdApps.Add "InfoPath",""
        dicConfigIdApps.Add "Skype for Business",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
        dicConfigIdApps.Add "Project",""
        dicConfigIdApps.Add "Visio",""
    Case "O365ProPlusRetail"
        dicConfigIdApps.Add "Access",""
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "Groove",""
        If iVM = 15 Then dicConfigIdApps.Add "InfoPath",""
        dicConfigIdApps.Add "Skype for Business",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
    Case "O365BusinessRetail"
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "Groove",""
        dicConfigIdApps.Add "Skype for Business",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
    Case "O365SmallBusPremRetail"
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "Groove",""
        dicConfigIdApps.Add "Skype for Business",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
    Case "VisioProRetail"
        dicConfigIdApps.Add "Visio",""
    Case "ProjectProRetail"
        dicConfigIdApps.Add "Project",""
    Case "AccessRetail"
        dicConfigIdApps.Add "Access",""
    Case "ExcelRetail"
        dicConfigIdApps.Add "Excel",""
    Case "GrooveRetail"
        dicConfigIdApps.Add "Groove",""
    Case "HomeBusinessRetail"
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Word",""
    Case "HomeStudentRetail"
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Word",""
    Case "InfoPathRetail"
        dicConfigIdApps.Add "InfoPath",""
    Case "LyncEntryRetail"
        dicConfigIdApps.Add "Skype for Business",""
    Case "LyncRetail"
        dicConfigIdApps.Add "Skype for Business",""
    Case "SkypeforBusinessEntryRetail"
        dicConfigIdApps.Add "Skype for Business",""
    Case "SkypeforBusinessRetail"
        dicConfigIdApps.Add "Skype for Business",""
    Case "ProfessionalRetail"
        dicConfigIdApps.Add "Access",""
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
    Case "O365HomePremRetail"
        dicConfigIdApps.Add "Access",""
        dicConfigIdApps.Add "Excel",""
        dicConfigIdApps.Add "OneNote",""
        dicConfigIdApps.Add "Outlook",""
        dicConfigIdApps.Add "PowerPoint",""
        dicConfigIdApps.Add "Publisher",""
        dicConfigIdApps.Add "Word",""
    Case "OneNoteRetail"
        dicConfigIdApps.Add "OneNote",""
    Case "OutlookRetail"
        dicConfigIdApps.Add "Outlook",""
    Case "PowerPointRetail"
        dicConfigIdApps.Add "PowerPoint",""
    Case "ProjectStdRetail"
        dicConfigIdApps.Add "Project",""
    Case "PublisherRetail"
        dicConfigIdApps.Add "Publisher",""
    Case "VisioStdRetail"
        dicConfigIdApps.Add "Visio",""
    Case "WordRetail"
        dicConfigIdApps.Add "Word",""
    End Select

End Sub

'-------------------------------------------------------------------------------
'   GetC2RVersionActive
'
'   Obtain the active C2R version
'   Returns an array with version numbers that have a C2R product
'   An active version is detected if the registry entry 'x-none' is found at 
'   HKLM\SOFTWARE\Microsoft\Office\XX.y\ClickToRun\ProductReleaseIDs\Active\culture
'-------------------------------------------------------------------------------
Sub GetC2RVersionActive()
    Dim key, sValue, sActiveConfiguration
    Dim arrKeys

    On Error Resume Next
    
    fv2SxS = False
    If RegEnumKey (HKLM, "Software\Microsoft\Office", arrKeys) Then
        For Each key in arrKeys
            Select Case LCase(key)
            Case "15.0"
                If RegReadValue (HKLM, "Software\Microsoft\Office\" & key & "\ClickToRun\ProductReleaseIDs\Active\culture","x-none", sValue, "REG_SZ") Then
                    AddToDic dicActiveC2RVersions, sValue, Left(sValue, 2)
                End If
            Case "clicktorun"
		        If RegReadValue(HKLM, "Software\Microsoft\Office\" & key & "\ProductReleaseIDs","ActiveConfiguration", sActiveConfiguration, "REG_SZ") Then
		        	If RegReadValue(HKLM, "Software\Microsoft\Office\" & key & "\ProductReleaseIDs\" & sActiveConfiguration & "\culture\x-none.15","Version", sValue, "REG_SZ") Then
                        fv2SxS = True
	                    AddToDic dicActiveC2RVersions, sValue, Left(sValue, 2)
                    End If
		        	If RegReadValue(HKLM, "Software\Microsoft\Office\" & key & "\ProductReleaseIDs\" & sActiveConfiguration & "\culture\x-none.16","Version", sValue, "REG_SZ") Then
	                    AddToDic dicActiveC2RVersions, sValue, Left(sValue, 2)
                    End If
                End If
            End Select
        Next
    End If
End Sub 'GetC2RVersionActive

'-------------------------------------------------------------------------------
'   GetBr
'
'-------------------------------------------------------------------------------
Function GetBr(sFfn)
    Dim sCh0, sCh1, sCh2, sCh3, sCh4, sCh5, sCh6, sCh7
    Dim sCh8, sCh9, sCh10, sCh11, sCh12, sCh13, sValue, sBr

    On Error Resume Next
    sValue = sFfn
    If Right(sValue, 1) = "/" Then sValue = Left(sValue, Len(sValue) - 1)
    If InStrRev(sValue, "/") > 0 Then sValue = Mid(sValue, InStrRev(sValue, "/") + 1)
    If Len(sValue) > 8 Then sValue = UCase(Left(sValue, 8))

    sBr = hAtS(Array("43","75","73","74","6F","6D"))
    sCh0 = hAtS(Array("33","39","31","36","38","44","37","45"))
    sCh1 = hAtS(Array("34","39","32","33","35","30","46","36"))
    sCh2 = hAtS(Array("36","34","32","35","36","41","46","45"))
    sCh3 = hAtS(Array("42","38","46","39","42","38","35","30"))
    sCh4 = hAtS(Array("37","46","46","42","43","36","42","46"))
    sCh5 = hAtS(Array("35","34","34","30","46","44","31","46"))
    sCh6 = hAtS(Array("46","32","45","37","32","34","43","31"))
    sCh7 = hAtS(Array("42","36","31","32","38","35","44","44"))
    sCh8 = hAtS(Array("31","44","32","44","32","45","41","36"))
    sCh9 = hAtS(Array("32","45","31","34","38","44","45","39"))
    sCh10 = hAtS(Array("35","34","36","32","45","45","45","35"))
    sCh11 = hAtS(Array("39","41","33","42","37","46","46","32"))
    sCh12 = hAtS(Array("46","34","46","30","32","34","43","38"))
    sCh13 = hAtS(Array("45","41","34","41","34","30","39","30"))

    If InStr(sValue, sCh0) > 0 Then sBr = hAtS(Array("4F","31","35","20","50","55"))
    If InStr(sValue, sCh1) > 0 Then sBr = hAtS(Array("4D","6F","6E","74","68","6C","79"))
    If InStr(sValue, sCh2) > 0 Then sBr = hAtS(Array("4D","6F","6E","74","68","6C","79","20","54","61","72","67","65","74","65","64","20","28","4D","54","29"))
    If InStr(sValue, sCh3) > 0 Then sBr = hAtS(Array("53","65","6D","69","2D","41","6E","6E","75","61","6C","20","54","61","72","67","65","74","65","64","20","28","53","41","54","29"))
    If InStr(sValue, sCh4) > 0 Then sBr = hAtS(Array("53","65","6D","69","2D","41","6E","6E","75","61","6C","20","28","53","41","29"))
    If InStr(sValue, sCh5) > 0 Then sBr = hAtS(Array("49","6E","73","69","64","65","72"))
    If InStr(sValue, sCh6) > 0 Then sBr = hAtS(Array("4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
    If InStr(sValue, sCh7) > 0 Then sBr = hAtS(Array("4D","53","49","54","20","45","6C","69","74","65"))
    If InStr(sValue, sCh8) > 0 Then sBr = hAtS(Array("4D","53","49","54","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
    If InStr(sValue, sCh9) > 0 Then sBr = hAtS(Array("49","6E","73","69","64","65","72","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
    If InStr(sValue, sCh10) > 0 Then sBr = hAtS(Array("49","6E","73","69","64","65","72","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
    If InStr(sValue, sCh11) > 0 Then sBr = hAtS(Array("4D","53","49","54","20","53","41","54"))
    If InStr(sValue, sCh12) > 0 Then sBr = hAtS(Array("4D","53","49","54","20","53","41"))
    If InStr(sValue, sCh13) > 0 Then sBr = hAtS(Array("4D","53","49","54","20","44","6F","67","66","6F","6F","64"))
    GetBr = sBr
End Function

'-------------------------------------------------------------------------------
'   AddBrDet
'
'-------------------------------------------------------------------------------
Sub AddBrDet(sFfn, dic)
    Dim sCh0, sCh1, sCh2, sCh3, sCh4, sCh5, sCh6, sCh7, sCh8, sCh9, sCh10, sCh11, sCh12, sCh13
    Dim sChN, sCh, sAu, sValue

    On Error Resume Next
    sValue = sFfn
    If Right(sValue, 1) = "/" Then sValue = Left(sValue, Len(sValue) - 1)
    If InStrRev(sValue, "/") > 0 Then sValue = Mid(sValue, InStrRev(sValue, "/") + 1)
    sFfn = sValue
    If Len(sFfn) > 8 Then sValue = UCase(Left(sFfn, 8))

    sChN = hAtS(Array("43","75","73","74","6F","6D"))
    sCh  = hAtS(Array("43","43"))
    sAu  = hAtS(Array("50","72","6F","64","75","63","74","69","6F","6E"))
    
    sCh0 = hAtS(Array("33","39","31","36","38","44","37","45"))
    sCh1 = hAtS(Array("34","39","32","33","35","30","46","36"))
    sCh2 = hAtS(Array("36","34","32","35","36","41","46","45"))
    sCh3 = hAtS(Array("42","38","46","39","42","38","35","30"))
    sCh4 = hAtS(Array("37","46","46","42","43","36","42","46"))
    sCh5 = hAtS(Array("35","34","34","30","46","44","31","46"))
    sCh6 = hAtS(Array("46","32","45","37","32","34","43","31"))
    sCh7 = hAtS(Array("42","36","31","32","38","35","44","44"))
    sCh8 = hAtS(Array("31","44","32","44","32","45","41","36"))
    sCh9 = hAtS(Array("32","45","31","34","38","44","45","39"))
    sCh10 = hAtS(Array("35","34","36","32","45","45","45","35"))
    sCh11 = hAtS(Array("39","41","33","42","37","46","46","32"))
    sCh12 = hAtS(Array("46","34","46","30","32","34","43","38"))
    sCh13 = hAtS(Array("45","41","34","41","34","30","39","30"))

    If InStr(sValue, sCh0) > 0 Then sChN = hAtS(Array("4F","31","35","20","50","55"))
    If InStr(sValue, sCh1) > 0 Then sChN = hAtS(Array("4D","6F","6E","74","68","6C","79"))
    If InStr(sValue, sCh2) > 0 Then
        sChN = hAtS(Array("4D","6F","6E","74","68","6C","79","20","54","61","72","67","65","74","65","64","20","28","4D","54","29"))
        sAu  = hAtS(Array("49","6E","73","69","64","65","72","73"))
    End If
    If InStr(sValue, sCh3) > 0 Then
        sChN = hAtS(Array("53","65","6D","69","2D","41","6E","6E","75","61","6C","20","54","61","72","67","65","74","65","64","20","28","53","41","54","29"))
        sCh  = hAtS(Array("46","52","44","43"))
        sAu  = hAtS(Array("49","6E","73","69","64","65","72","73"))
    End If
    If InStr(sValue, sCh4) > 0 Then
        sChN = hAtS(Array("53","65","6D","69","2D","41","6E","6E","75","61","6C","20","28","53","41","29"))
        sCh  = hAtS(Array("44","43"))
    End If
    If InStr(sValue, sCh5) > 0 Then
        sChN = hAtS(Array("49","6E","73","69","64","65","72"))
        sCh  = hAtS(Array("44","65","76","4D","61","69","6E"))
        sAu  = hAtS(Array("49","6E","73","69","64","65","72","73"))
    End If
    If InStr(sValue, sCh6) > 0 Then
        sChN = hAtS(Array("4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
        sCh  = hAtS(Array("4C","54","53","43"))
    End If
    If InStr(sValue, sCh7) > 0 Then
        sChN = hAtS(Array("4D","53","49","54","20","45","6C","69","74","65"))
        sCh  = hAtS(Array("44","65","76","4D","61","69","6E"))
        sAu  = hAtS(Array("4D","69","63","72","6F","73","6F","66","74"))
    End If
    If InStr(sValue, sCh8) > 0 Then
        sChN = hAtS(Array("4D","53","49","54","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
        sCh  = hAtS(Array("4C","54","53","43"))
        sAu  = hAtS(Array("4D","69","63","72","6F","73","6F","66","74"))
    End If
    If InStr(sValue, sCh9) > 0 Then
        sChN = hAtS(Array("49","6E","73","69","64","65","72","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
        sCh  = hAtS(Array("4C","54","53","43"))
        sAu  = hAtS(Array("49","6E","73","69","64","65","72","73"))
    End If
    If InStr(sValue, sCh10) > 0 Then
        sChN = hAtS(Array("49","6E","73","69","64","65","72","20","4F","32","30","31","39","20","56","6F","6C","75","6D","65"))
        sAu  = hAtS(Array("4D","69","63","72","6F","73","6F","66","74"))
    End If
    If InStr(sValue, sCh11) > 0 Then
        sChN = hAtS(Array("4D","53","49","54","20","53","41","54"))
        sCh  = hAtS(Array("46","52","44","43"))
        sAu  = hAtS(Array("4D","69","63","72","6F","73","6F","66","74"))
    End If
    If InStr(sValue, sCh12) > 0 Then
        sChN = hAtS(Array("4D","53","49","54","20","53","41"))
        sCh  = hAtS(Array("44","43"))
        sAu  = hAtS(Array("4D","69","63","72","6F","73","6F","66","74"))
    End If
    If InStr(sValue, sCh13) > 0 Then
        sChN = hAtS(Array("4D","53","49","54","20","44","6F","67","66","6F","6F","64"))
        sCh  = hAtS(Array("44","65","76","4D","61","69","6E"))
        sAu  = hAtS(Array("44","6F","67","66","6F","6F","64"))
    End If

    AddToDic dic, hAtS(Array("43","44","4E","20","4E","61","6D","65")), sChN
    AddToDic dic, hAtS(Array("43","68","61","6E","6E","65","6C")), sCh
    AddToDic dic, hAtS(Array("41","75","64","69","65","6E","63","65")), sAu
    AddToDic dic, hAtS(Array("46","46","4E")), sFfn
End Sub 'GetCh

'-------------------------------------------------------------------------------
'   GetKeyComponentStates
'
'   Obtain the key component states for a product
'   Returns a string with the applicable states for the key components
'   Application: Name, ExeName, VersionMajor, Version, InstallState, 
'                InstallStateString, ComponentId, FeatureName, Path 
'-------------------------------------------------------------------------------
Function GetKeyComponentStates(sProductCode, fVirtualized)
    Dim sSkuId, sReturn, sComponents, sProd, sName, sExeName, sBitness
    Dim sVersion, sInstallState, sInstallStateString, sPath
    Dim component
    Dim sFeatureName
    Dim iVM
    Dim fIsWW
    Dim arrComponents

    On Error Resume Next
    sReturn = ""
    GetKeyComponentStates = sReturn
	If IsOfficeProduct(sProductCode) Then
        iVM = GetVersionMajor(sProductCode)
        If iVM < 12 Then 
            sSkuId = Mid(sProductCode, 4, 2)
            fIsWW = True
        Else
            sSkuId = Mid(sProductCode, 11, 4)
            fIsWW = (Mid(sProductCode, 16, 4) = "0000")
        End If
        If sProductCode = sPackageGuid Then
            fIsWW = True
            If iVMVirtOverride > 0 Then
                iVM = iVMVirtOverride
            Else
                If dicC2RPropV2.Count > 0 Then iVM = 15
                If dicC2RPropV3.Count > 0 Then iVM = 16
            End If
        End If
    End If
    If NOT fIsWW Then Exit Function 
	
    sComponents = GetCoreComponentCode(sSkuId, iVM, fVirtualized)
    arrComponents = Split(sComponents, ";")
    sProd = sProductCode
    If fVirtualized Then sProd = UCase(dicProductCodeC2R.Item(iVM))

    ' get the component data registered to the product
    For Each component in arrComponents
        If dicKeyComponents.Exists (component) Then
            If InStr (UCase(dicKeyComponents.Item (component)), UCase(sProd)) > 0 Then
                sFeatureName = GetMsiFeatureName (component)
                sName = GetApplicationName (sFeatureName)
                sBitness = GetComponentBitness (component)
                sExeName = GetApplicationExeName (component)
                sPath = oMsi.ComponentPath (sProd, component)
                If oFso.FileExists (sPath) Then sVersion = oFso.GetFileVersion (sPath) Else sVersion = ""
                sInstallState = oMsi.FeatureState (sProd, sFeatureName)
                sInstallStateString = TranslateFeatureState (sInstallState)
                sReturn = sReturn & sName & "_" & sBitness & "," & sExeName & "," & iVM & "," & sVersion & "," & sInstallState & "," & sInstallStateString & "," & component & "," & sFeatureName & "," & sPath & ";"
            End If
        End If
    Next 'component
    If NOT sReturn = "" Then sReturn = Left(sReturn, Len(sReturn) - 1)
    GetKeyComponentStates = sReturn

End Function 'GetKeyComponentStates

'-------------------------------------------------------------------------------
'   GetCoreComponentCode
'
'   Returns the component code(s) for core application .exe files
'   from the SKU element of the productcode
'-------------------------------------------------------------------------------
Function GetCoreComponentCode (sSkuId, iVM, fVirtualized)
    Dim sReturn
    Dim arrTmp

    On Error Resume Next
    sReturn = ""
    GetCoreComponentCode = sReturn
    Select Case iVM
    Case 16
        sReturn = Join(Array(CID_ACC16_64, CID_ACC16_32, CID_XL16_64, CID_XL16_32, CID_GRV16_64, CID_GRV16_32, CID_LYN16_64, CID_LYN16_32, CID_ONE16_64, CID_ONE16_32, CID_OL16_64, CID_OL16_32, CID_PPT16_64, CID_PPT16_32, CID_PRJ16_64, CID_PRJ16_32, CID_PUB16_64, CID_PUB16_32, CID_IP16_64, CID_VIS16_64, CID_VIS16_32, CID_WD16_64, CID_WD16_32, CID_SPD16_64, CID_SPD16_32, CID_MSO16_64, CID_MSO16_32), ";")
        If fVirtualized Then
            Select Case sSkuId
            Case "0015" : sReturn = CID_ACC16_64 & ";" & CID_ACC16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Access
            Case "0016","0029" : sReturn = CID_XL16_64 & ";" & CID_XL16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Excel
            Case "0018" : sReturn = CID_PPT16_64 & ";" & CID_PPT16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'PowerPoint
            Case "0019" : sReturn = CID_PUB16_64 & ";" & CID_PUB16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Publisher
            Case "001A","00E0" : sReturn = CID_OL16_64 & ";" & CID_OL16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Outlook
            Case "001B","002B" : sReturn = CID_WD16_64 & ";" & CID_WD16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Word
            Case "0027" : sReturn = CID_PRJ16_64 & ";" & CID_PRJ16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Project
            Case "0044" : sReturn = CID_IP16_64 & ";"  & CID_MSO16_64 & ";" & CID_MSO16_32 'InfoPath
            Case "0017" : sReturn = CID_SPD16_64 & ";" & CID_SPD16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'SharePointDesigner
            Case "0051","0053","0057" : sReturn = CID_VIS16_64 & ";" & CID_VIS16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Visio
            Case "00A1","00A3" : sReturn = CID_ONE16_64 & ";" & CID_ONE16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'OneNote
            Case "00BA" : sReturn = CID_GRV16_64 & ";" & CID_GRV16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Groove
            Case "012B","012C" : sReturn = CID_LYN16_64 & ";" & CID_LYN16_32 & ";" & CID_MSO16_64 & ";" & CID_MSO16_32 'Lync
            Case Else : sReturn = ""
            End Select
        End If
    Case 15
        sReturn = Join(Array(CID_ACC15_64, CID_ACC15_32, CID_XL15_64, CID_XL15_32, CID_GRV15_64, CID_GRV15_32, CID_LYN15_64, CID_LYN15_32, CID_ONE15_64, CID_ONE15_32, CID_OL15_64, CID_OL15_32, CID_PPT15_64, CID_PPT15_32, CID_PRJ15_64, CID_PRJ15_32, CID_PUB15_64, CID_PUB15_32, CID_IP15_64, CID_IP15_32, CID_VIS15_64, CID_VIS15_32, CID_WD15_64, CID_WD15_32, CID_SPD15_64, CID_SPD15_32, CID_MSO15_64, CID_MSO15_32), ";")
        If fVirtualized Then
            Select Case sSkuId
            Case "0015" : sReturn = CID_ACC15_64 & ";" & CID_ACC15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Access
            Case "0016","0029" : sReturn = CID_XL15_64 & ";" & CID_XL15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Excel
            Case "0018" : sReturn = CID_PPT15_64 & ";" & CID_PPT15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'PowerPoint
            Case "0019" : sReturn = CID_PUB15_64 & ";" & CID_PUB15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Publisher
            Case "001A","00E0" : sReturn = CID_OL15_64 & ";" & CID_OL15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Outlook
            Case "001B","002B" : sReturn = CID_WD15_64 & ";" & CID_WD15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Word
            Case "0027" : sReturn = CID_PRJ15_64 & ";" & CID_PRJ15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Project
            Case "0044" : sReturn = CID_IP15_64 & ";" & CID_IP15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'InfoPath
            Case "0017" : sReturn = CID_SPD15_64 & ";" & CID_SPD15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'SharePointDesigner
            Case "0051","0053","0057" : sReturn = CID_VIS15_64 & ";" & CID_VIS15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Visio
            Case "00A1","00A3" : sReturn = CID_ONE15_64 & ";" & CID_ONE15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'OneNote
            Case "00BA" : sReturn = CID_GRV15_64 & ";" & CID_GRV15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Groove
            Case "012B","012C" : sReturn = CID_LYN15_64 & ";" & CID_LYN15_32 & ";" & CID_MSO15_64 & ";" & CID_MSO15_32 'Lync
            Case Else : sReturn = ""
            End Select
        End If
    Case 14
        sReturn = Join(Array(CID_ACC14_64, CID_ACC14_32, CID_XL14_64, CID_XL14_32, CID_GRV14_64, CID_GRV14_32, CID_ONE14_64, CID_ONE14_32, CID_OL14_64, CID_OL14_32, CID_PPT14_64, CID_PPT14_32, CID_PRJ14_64, CID_PRJ14_32, CID_PUB14_64, CID_PUB14_32, CID_IP14_64, CID_IP14_32, CID_VIS14_64, CID_VIS14_32, CID_WD14_64, CID_WD14_32, CID_SPD14_64, CID_SPD14_32, CID_MSO14_64, CID_MSO14_32), ";")
    Case 12
        sReturn = Join(Array(CID_ACC12, CID_XL12, CID_GRV12, CID_ONE12, CID_OL12, CID_PPT12, CID_PRJ12, CID_PUB12, CID_IP12, CID_VIS12, CID_WD12, CID_SPD12, CID_MSO12), ";")
    Case 11
        sReturn = Join(Array(CID_ACC11, CID_XL11, CID_ONE11, CID_OL11, CID_PPT11, CID_PRJ11, CID_PUB11, CID_IP11, CID_VIS11, CID_WD11, CID_SPD11, CID_MSO11), ";")
    End Select
    GetCoreComponentCode = sReturn
End Function 'GetCoreComponentCode

'-------------------------------------------------------------------------------
'   GetMsiFeatureName
'
'   Get the known Feature name for a core component
'-------------------------------------------------------------------------------
Function GetMsiFeatureName (sComponentId)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sComponentId
    Case CID_ACC16_64, CID_ACC16_32, CID_ACC15_64, CID_ACC15_32, CID_ACC14_64, CID_ACC14_32, CID_ACC12, CID_ACC11
        sReturn = "ACCESSFiles"
    Case CID_XL16_64, CID_XL16_32, CID_XL15_64, CID_XL15_32, CID_XL14_64, CID_XL14_32, CID_XL12, CID_XL11
        sReturn = "EXCELFiles"
    Case CID_GRV16_64, CID_GRV16_32, CID_GRV15_64, CID_GRV15_32
        sReturn = "GrooveFiles2"
    Case CID_GRV14_64, CID_GRV14_32, CID_GRV12
        sReturn = "GrooveFiles"
    Case CID_LYN16_64, CID_LYN16_32, CID_LYN15_64, CID_LYN15_32
        sReturn = "Lync_CoreFiles"
    Case CID_MSO16_64, CID_MSO16_32, CID_MSO15_64, CID_MSO15_32, CID_MSO14_64, CID_MSO14_32, CID_MSO12, CID_MSO11
        sReturn = "ProductFiles"
    Case CID_ONE16_64, CID_ONE16_32, CID_ONE15_64, CID_ONE15_32, CID_ONE14_64, CID_ONE14_32, CID_ONE12, CID_ONE11
        sReturn = "OneNoteFiles"
    Case CID_OL16_64, CID_OL16_32, CID_OL15_64, CID_OL15_32, CID_OL14_64, CID_OL14_32, CID_OL12, CID_OL11
        sReturn = "OUTLOOKFiles"
    Case CID_PPT16_64, CID_PPT16_32, CID_PPT15_64, CID_PPT15_32, CID_PPT14_64, CID_PPT14_32, CID_PPT12, CID_PPT11
        sReturn = "PPTFiles"
    Case CID_PRJ16_64, CID_PRJ16_32, CID_PRJ15_64, CID_PRJ15_32, CID_PRJ14_64, CID_PRJ14_32, CID_PRJ12, CID_PRJ11
        sReturn = "PROJECTFiles"
    Case CID_PUB16_64, CID_PUB16_32, CID_PUB15_64, CID_PUB15_32, CID_PUB14_64, CID_PUB14_32, CID_PUB12, CID_PUB11
        sReturn = "PubPrimary"
    Case CID_IP16_64, CID_IP15_64, CID_IP15_32, CID_IP14_64, CID_IP14_32, CID_IP12, CID_IP11
        sReturn = "XDOCSFiles"
    Case CID_VIS16_64, CID_VIS16_32, CID_VIS15_64, CID_VIS15_32, CID_VIS14_64, CID_VIS14_32, CID_VIS12, CID_VIS11
        sReturn = "VisioCore"
    Case CID_WD16_64, CID_WD16_32, CID_WD15_64, CID_WD15_32, CID_WD14_64, CID_WD14_32, CID_WD12, CID_WD11
        sReturn = "WORDFiles"
    Case CID_SPD16_64, CID_SPD16_32, CID_SPD15_64, CID_SPD15_32, CID_SPD14_64, CID_SPD14_32, CID_SPD12
        sReturn = "WAC_CoreSPD"
    Case CID_SPD11
        sReturn = "FPClientFiles"
    End Select
    GetMsiFeatureName = sReturn
End Function 'GetMsiFeatureName

'-------------------------------------------------------------------------------
'   GetApplicationName
'
'   Get the friendly name from the FeatureName
'-------------------------------------------------------------------------------
Function GetApplicationName (sFeatureName)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sFeatureName
    Case "ACCESSFiles"                  : sReturn = "Access"
    Case "EXCELFiles"                   : sReturn = "Excel"
    Case "GrooveFiles2","GrooveFiles"  : sReturn = "Groove"
    Case "Lync_CoreFiles"               : sReturn = "Skype for Business"
    Case "OneNoteFiles"                 : sReturn = "OneNote"
    Case "OUTLOOKFiles"                 : sReturn = "Outlook"
    Case "PPTFiles"                     : sReturn = "PowerPoint"
    Case "ProductFiles"                 : sReturn = "Mso"
    Case "PROJECTFiles"                 : sReturn = "Project"
    Case "PubPrimary"                   : sReturn = "Publisher"
    Case "XDOCSFiles"                   : sReturn = "InfoPath"
    Case "VisioCore"                    : sReturn = "Visio"
    Case "WORDFiles"                    : sReturn = "Word"
    Case "WAC_CoreSPD","FPClientFiles" : sReturn = "SharePoint Designer"
    End Select
    GetApplicationName = sReturn
End Function 'GetApplicationName

'-------------------------------------------------------------------------------
'   GetApplicationExeName
'
'   Get the friendly name from the FeatureName
'-------------------------------------------------------------------------------
Function GetApplicationExeName (sComponentId)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sComponentId
    Case CID_ACC16_64, CID_ACC16_32, CID_ACC15_64, CID_ACC15_32, CID_ACC14_64, CID_ACC14_32, CID_ACC12, CID_ACC11
        sReturn = "MSACCESS.EXE"
    Case CID_XL16_64, CID_XL16_32, CID_XL15_64, CID_XL15_32, CID_XL14_64, CID_XL14_32, CID_XL12, CID_XL11
        sReturn = "EXCEL.EXE"
    Case CID_GRV16_64, CID_GRV16_32, CID_GRV15_64, CID_GRV15_32, CID_GRV14_64, CID_GRV14_32, CID_GRV12
        sReturn = "GROOVE.EXE"
    Case CID_LYN16_64, CID_LYN16_32, CID_LYN15_64, CID_LYN15_32
        sReturn = "LYNC.EXE"
    Case CID_MSO16_64, CID_MSO16_32, CID_MSO15_64, CID_MSO15_32, CID_MSO14_64, CID_MSO14_32, CID_MSO12, CID_MSO11
        sReturn = "MSO.DLL"
    Case CID_ONE16_64, CID_ONE16_32, CID_ONE15_64, CID_ONE15_32, CID_ONE14_64, CID_ONE14_32, CID_ONE12, CID_ONE11
        sReturn = "ONENOTE.EXE"
    Case CID_OL16_64, CID_OL16_32, CID_OL15_64, CID_OL15_32, CID_OL14_64, CID_OL14_32, CID_OL12, CID_OL11
        sReturn = "OUTLOOK.EXE"
    Case CID_PPT16_64, CID_PPT16_32, CID_PPT15_64, CID_PPT15_32, CID_PPT14_64, CID_PPT14_32, CID_PPT12, CID_PPT11
        sReturn = "POWERPNT.EXE"
    Case CID_PRJ16_64, CID_PRJ16_32, CID_PRJ15_64, CID_PRJ15_32, CID_PRJ14_64, CID_PRJ14_32, CID_PRJ12, CID_PRJ11
        sReturn = "WINPROJ.EXE"
    Case CID_PUB16_64, CID_PUB16_32, CID_PUB15_64, CID_PUB15_32, CID_PUB14_64, CID_PUB14_32, CID_PUB12, CID_PUB11
        sReturn = "MSPUB.EXE"
    Case CID_IP16_64, CID_IP15_64, CID_IP15_32, CID_IP14_64, CID_IP14_32, CID_IP12, CID_IP11
        sReturn = "INFOPATH.EXE"
    Case CID_VIS16_64, CID_VIS16_32, CID_VIS15_64, CID_VIS15_32, CID_VIS14_64, CID_VIS14_32, CID_VIS12, CID_VIS11
        sReturn = "VISIO.EXE"
    Case CID_WD16_64, CID_WD16_32, CID_WD15_64, CID_WD15_32, CID_WD14_64, CID_WD14_32, CID_WD12, CID_WD11
        sReturn = "WINWORD.EXE"
    Case CID_SPD16_64, CID_SPD16_32, CID_SPD15_64, CID_SPD15_32, CID_SPD14_64, CID_SPD14_32, CID_SPD12
        sReturn = "SPDESIGN.EXE"
    Case CID_SPD11
        sReturn = "FRONTPG.EXE"
    End Select
    GetApplicationExeName = sReturn
End Function 'GetApplicationExeName

'-------------------------------------------------------------------------------
'   GetComponentBitness
'
'   Get the bitness for the component (x86 or x64)
'-------------------------------------------------------------------------------
Function GetComponentBitness (sComponentId)
    Dim sReturn

    On Error Resume Next
    sReturn = ""
    Select Case sComponentId
    Case CID_SPD14_64, CID_SPD15_64, CID_SPD16_64, CID_WD14_64, CID_WD15_64, CID_WD16_64, CID_VIS14_64, CID_VIS15_64, CID_VIS16_64, CID_IP14_64, CID_IP15_64, CID_IP16_64, CID_PUB14_64, CID_PUB15_64, CID_PUB16_64, CID_PRJ14_64, CID_PRJ15_64, CID_PRJ16_64, CID_PPT14_64, CID_PPT15_64, CID_ACC16_64, CID_ACC15_64, CID_ACC14_64, CID_XL16_64, CID_XL15_64, CID_XL14_64, CID_GRV16_64, CID_GRV15_64, CID_GRV14_64, CID_LYN16_64, CID_LYN15_64, CID_MSO16_64, CID_MSO15_64, CID_MSO14_64, CID_ONE16_64, CID_ONE15_64, CID_ONE14_64, CID_OL16_64, CID_OL15_64, CID_OL14_64, CID_PPT16_64
        sReturn = "x64"
    Case CID_SPD11, CID_SPD14_32, CID_SPD12, CID_SPD15_32, CID_SPD16_32, CID_WD14_32, CID_WD12, CID_WD11, CID_WD15_32, CID_WD16_32, CID_VIS14_32, CID_VIS12, CID_VIS11, CID_VIS15_32, CID_VIS16_32, CID_IP14_32, CID_IP12, CID_IP11, CID_IP15_32, CID_PUB14_32, CID_PUB12, CID_PUB11, CID_PUB15_32, CID_PUB16_32, CID_PRJ14_32, CID_PRJ12, CID_PRJ11, CID_PRJ15_32, CID_PRJ16_32, CID_PPT14_32, CID_PPT12, CID_PPT11, CID_PPT15_32, CID_PPT16_32, CID_ACC16_32, CID_ACC15_32, CID_ACC14_32, CID_ACC12, CID_ACC11, CID_XL16_32, CID_XL15_32, CID_XL14_32, CID_XL12, CID_XL11, CID_GRV16_32, CID_GRV15_32, CID_GRV14_32, CID_GRV12, CID_LYN16_32, CID_LYN15_32, CID_MSO16_32, CID_MSO15_32, CID_MSO14_32, CID_MSO12, CID_MSO11, CID_ONE16_32, CID_ONE15_32, CID_ONE14_32, CID_ONE12, CID_ONE11, CID_OL16_32, CID_OL15_32, CID_OL14_32, CID_OL12, CID_OL11
        sReturn = "x86"
    End Select
    GetComponentBitness = sReturn
End Function 'GetComponentBitness

'-------------------------------------------------------------------------------
'   GetConfigName
'
'   Get the configuration name from the ARP key name
'-------------------------------------------------------------------------------
Function GetConfigName(ArpItem)
    Dim sCurKey, sValue, sDisplayVersion, sUninstallString
    dim sCulture, sConfigName
    Dim iLeft, iRight
    Dim fSystemComponent0, fDisplayVersion, fUninstallString

    sCurKey = REG_ARP & ArpItem & "\"
    sValue = ""
    sDisplayVersion = ""
    fSystemComponent0 = NOT (RegReadValue(HKLM, sCurKey, "SystemComponent", sValue, "REG_DWORD") AND (sValue = "1"))
    fDisplayVersion = RegReadValue(HKLM, sCurKey, "DisplayVersion", sValue, "REG_SZ")
    If fDisplayVersion Then
        sDisplayVersion = sValue
        If Len(sValue) > 1 Then
            fDisplayVersion = (Left(sValue, 2) = "15")
        Else
            fDisplayVersion = False
        End If
    End If
    fUninstallString = RegReadValue(HKLM, sCurKey, "UninstallString", sUninstallString, "REG_SZ")

    'C2R
    If (fSystemComponent0 AND fDisplayVersion AND InStr(UCase(sUninstallString), UCase(OREGREFC2R15)) > 0) Then
        iLeft = InStr(ArpItem, " - ") + 2
        iRight = InStr(iLeft, ArpItem, " - ") - 1
        If iRight > 0 Then
            sConfigName = Trim(Mid(ArpItem, iLeft, (iRight - iLeft)))
            sCulture = Mid(ArpItem, iRight + 3)
        Else
            sConfigName = Trim(Left(ArpItem, iLeft - 3))
            sCulture = Mid(ArpItem, iLeft)
        End If
        sConfigName = Replace(sConfigName, "Microsoft","")
        sConfigName = Replace(sConfigName, "Office","")
        sConfigName = Replace(sConfigName, "Professional","Pro")
        sConfigName = Replace(sConfigName, "Standard","Std")
        sConfigName = Replace(sConfigName, "(Technical Preview)","")
        sConfigName = Replace(sConfigName, "15","")
        sConfigName = Replace(sConfigName, "2013","")
        sConfigName = Replace(sConfigName, " ","")
        'sConfigName = Replace(sConfigName, "Project","Prj")
        'sConfigName = Replace(sConfigName, "Visio","Vis")
        GetConfigName = sConfigName
        Exit Function
    End If

    'Standalone helper MSI products
    Select Case Mid(ArpItem, 11, 4)
    Case "007E","008F","008C","00DD" 'C2R Integration Components
        GetConfigName = "Habanero"
    Case "24E1","237A"
        GetConfigName = "MSOIDLOGIN"
    Case Else
        GetConfigName = ""
    End Select

End Function 'GetConfigName

'=======================================================================================================

Sub CopyToMaster (Arr)
    Dim i, j, n
    On Error Resume Next
    
    For n = 0 To UBound(arrMaster)
        If (IsEmpty(arrMaster(n, COL_PRODUCTCODE)) OR Not (Len(arrMaster(n, COL_PRODUCTCODE)) > 0)) Then Exit For
    Next 'n
    
    For i = 0 To UBound(Arr, 1)
        For j = 0 To UBound(Arr, 2)
            arrMaster(i+n, j) = arr(i, j)
        Next 'j
    Next 'i
End Sub
'=======================================================================================================

Function GetObjProductState(ProdX, iContext, sSid, iSource)
    Dim iState
    On Error Resume Next

    Select Case iSource
    Case 0 'Registry
        iState = GetRegProductState(ProdX, iContext, sSid)
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        iState = ProdX.State
    Case Else
        iState = -1
    End Select
    GetObjProductState = iState
End Function

'=======================================================================================================
Function GetPrimaryProductId(sProductCode)
    Dim sPrimProdId
    Dim iState
    On Error Resume Next

    If InStr(sProductCode, "0000000FF1CE") > 0 Then sPrimProdId = Mid(sProductCode, 11, 4) Else sPrimProdId = Mid(sProductCode, 4, 2)
    GetPrimaryProductId = sPrimProdId
End Function
'=======================================================================================================

Function TranslateObjProductState(iState)
    Dim sState
    On Error Resume Next

    Select Case iState

    Case INSTALLSTATE_ADVERTISED '1
        sState = "Advertised"
    Case INSTALLSTATE_ABSENT '2
        sState = "Absent"
    Case INSTALLSTATE_DEFAULT '5
        sState = "Installed"    
    Case INSTALLSTATE_VIRTUALIZED '8
        sState = "Virtualized"  
    Case Else '- 1
        sState = "Unknown"
    End Select
    TranslateObjProductState = sState
End Function
'=======================================================================================================

'Obtain the ProductName for product
Function GetObjProductName(ProdX, sProductCode, iContext, sSid, iSource)
    Dim sName
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        sName = GetRegProductName(ProdX, iContext, sSid)
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        Err.Clear
        sName = ProdX.InstallProperty ("ProductName")
        If Not Err = 0 Then
            sName = GetRegProductName(GetCompressedGuid(sProductCode), iContext, sSid)
        End If 'Err

    Case Else

    End Select
    GetObjProductName = sName
End Function
'=======================================================================================================

'Get the current products UserSID.
Function GetObjUserSid(ProdX, sSid, iSource)
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        sSid = sSid
    Case 2 'WI 2.x

    Case 3 'WI >=3.x
        sSid = ProdX.UserSid
    Case Else

    End Select
    
    GetObjUserSid = sSid
End Function
'=======================================================================================================

Function GetObjContext(ProdX, iContext, iSource)
    On Error Resume Next

    If iSource = 3 Then iContext = ProdX.Context
    GetObjContext = iContext
End Function
'=======================================================================================================

Function TranslateObjContext(iContext)
    Dim sContext
    On Error Resume Next

    Select Case iContext

    Case MSIINSTALLCONTEXT_USERMANAGED '1
        sContext = "User Managed"
    Case MSIINSTALLCONTEXT_USERUNMANAGED '2
        sContext = "User Unmanaged"
    Case MSIINSTALLCONTEXT_MACHINE '4
        sContext = "Machine"
    Case MSIINSTALLCONTEXT_ALL '7
        sContext = "All"
    Case MSIINSTALLCONTEXT_C2RV2 '8
        sContext = "Machine C2Rv2"
    Case MSIINSTALLCONTEXT_C2RV3 '15
        sContext = "Machine C2Rv3"
    Case Else
        sContext = "Unknown"

    End Select
    TranslateObjContext = sContext
End Function
'=======================================================================================================

Function GetObjGuid(ProdX, iSource)
    Dim sGuid
    On Error Resume Next

    Select Case iSource

    Case 0 'Registry
        If (IsValidGuid(ProdX, GUID_COMPRESSED) OR fGuidCaseWarningOnly) Then 
            sGuid = GetExpandedGuid(ProdX)
        Else
            sGuid = ProdX
        End If

    Case 2 'WI 2.x
        sGuid = ProdX

    Case 3 'WI >=3.x
        sGuid = ProdX.ProductCode

    Case Else

    End Select
    GetObjGuid = sGuid
End Function
'=======================================================================================================

Sub WritePLArrayEx(iSource, Arr, Obj, iContext, sSid)
    Dim ProdX, Product, sProductCode, sProductName
    Dim i, n, iDimCnt, iPosUMP
    On Error Resume Next
    
    If CheckObject(Obj) Or CheckArray(Obj) Then
        i = 0
        If CheckObject(Obj) Then 
            ReDim Arr(Obj.Count - 1, UBOUND_MASTER)
        Else
            If UBound(Obj)>UBound(Arr) Then ReDim Arr(UBound(Obj), UBOUND_MASTER)
            Do While Not IsEmpty(Arr(i, 0))
                i = i + 1
                If i >= UBound(Obj) Then Exit Do
            Loop
        End If 'CheckObject
        For Each ProdX in Obj
        ' preset Recordset with Default Error
            For n = 0 to 4
                Arr(i, n) = "Preset Error String"
            Next 'n
        ' ProductCode
            Arr(i, COL_PRODUCTCODE) = GetObjGUID(ProdX, iSource)
        ' PrimaryProductId
            Arr(i, COL_PRIMARYPRODUCTID) = GetPrimaryProductId(Arr(i, COL_PRODUCTCODE))
        ' ProductContext
            Arr(i, COL_CONTEXT) = GetObjContext(ProdX, iContext, iSource)
            Arr(i, COL_CONTEXTSTRING) = TranslateObjContext(Arr(i, COL_CONTEXT))
        ' SID    
            Arr(i, COL_USERSID) = GetObjUserSid(ProdX, sSid, iSource)
        ' ProductName
            Arr(i, COL_PRODUCTNAME) = GetObjProductName(ProdX, Arr(i, COL_PRODUCTCODE), Arr(i, COL_CONTEXT), Arr(i, COL_USERSID), iSource)
        ' ProductState    
            Arr(i, COL_STATE) = GetObjProductState(ProdX, Arr(i, COL_CONTEXT), Arr(i, COL_USERSID), iSource)
            Arr(i, COL_STATESTRING) = TranslateObjProductState(Arr(i, COL_STATE))
        ' write to cache
            CacheLog LOGPOS_RAW, LOGHEADING_NONE, Null, Arr(i, COL_PRODUCTCODE) & "," & Arr(i, COL_CONTEXT) & "," & Arr(i, COL_USERSID) & "," & _
                Arr(i, COL_PRODUCTNAME) & "," & Arr(i, COL_STATE) 
        ' ARP ProductName    
            Arr(i, COL_ARPPRODUCTNAME) = GetArpProductname(Arr(i, COL_PRODUCTCODE))
        ' Guid validation
            If Not IsValidGuid(Arr(i, COL_PRODUCTCODE), GUID_UNCOMPRESSED) Then 
                If fGuidCaseWarningOnly Then 
                    Arr(i, COL_NOTES) = Arr(i, COL_NOTES) & ERR_CATEGORYNOTE & ERR_GUIDCASE & CSV
                Else
                    Arr(i, COL_ERROR) = Arr(i, COL_ERROR) & ERR_CATEGORYERROR & sError & CSV
                    Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, sError & DSV & Arr(i, COL_PRODUCTCODE) & DSV & _
                    Arr(i, COL_PRODUCTNAME)& DSV & BPA_GUID
                End If
            End If
        ' Virtual flag
            Arr(i, COL_VIRTUALIZED) = 0
            If iContext = MSIINSTALLCONTEXT_C2RV2 Then Arr(i, COL_VIRTUALIZED) = 1
            If iContext = MSIINSTALLCONTEXT_C2RV3 Then Arr(i, COL_VIRTUALIZED) = 1
        ' InstallType flag
            Arr(i, COL_INSTALLTYPE) = "MSI"
            If iContext = MSIINSTALLCONTEXT_C2RV2 Then Arr(i, COL_INSTALLTYPE) = "C2R"
            If iContext = MSIINSTALLCONTEXT_C2RV3 Then Arr(i, COL_INSTALLTYPE) = "C2R"

            i = i + 1
        Next 'ProdX
    End If 'ObjectList.Count > 0
End Sub 'WritePLArrayEx
'=======================================================================================================

Sub InitPLArrays
    On Error Resume Next
    
    ReDim arrAllProducts(-1)
    ReDim arrVirtProducts(-1)
    ReDim arrVirt2Products(-1)
    ReDim arrVirt3Products(-1)
    ReDim arrMProducts(-1)
    ReDim arrMVProducts(-1)
    ReDim arrUUProducts(-1)
    ReDim arrUMProducts(-1)
End Sub

'=======================================================================================================
'Module Prerequisites
'=======================================================================================================
Function CheckPreReq ()
    On Error Resume Next
    Dim sActiveSub, sErrHnd
    Dim sPreReqError, sDebugLogName, sSubKeyName
    Dim i, iFoo
    Dim hDefKey, lAccPermLevel
    Dim oCompItem, oItem
    sActiveSub = "CheckPreReq" : sErrHnd = "_ErrorHandler"

    fIsCriticalError = False
    CheckPreReq = True : fIsAdmin = True : fIsElevated = True
    sPreReqError = vbNullString
    Err.Clear

    'Create the WScript Shell Object
    Set oShell = CreateObject("WScript.Shell"): CheckError sActiveSub, sErrHnd
    sComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%"): CheckError sActiveSub, sErrHnd
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%"): CheckError sActiveSub, sErrHnd

    'Create the Windows Installer Object
    Set oMsi = CreateObject("WindowsInstaller.Installer"): CheckError sActiveSub, sErrHnd
    iWiVersionMajor = Left(oMsi.Version, Instr(oMsi.Version, ".") - 1)
    If (CheckPreReq = True And iWiVersionMajor < 2) Then CheckPreReq = False

    'Create the FileSystemObject
    Set oFso = CreateObject("Scripting.FileSystemObject"): CheckError sActiveSub, sErrHnd

    'Connect to WMI Registry Provider
    Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv"): CheckError sActiveSub, sErrHnd
    
    Set oWmiLocal = GetObject("winmgmts:\\.\root\cimv2")
    Set oCompItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
    For Each oItem In oCompItem
        sSystemType = oItem.SystemType
        f64 = Instr(Left(oItem.SystemType, 3), "64") > 0
    Next
    
    'Check registry access permissions
    'Failure will not terminate the scipt but noted in the log
    Set dicActiveC2RVersions = CreateObject("Scripting.Dictionary")
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = "SOFTWARE\Microsoft\Windows"
    For i = 1 To 4
        Select Case i
        Case 1 : lAccPermLevel = KEY_QUERY_VALUE
        Case 2 : lAccPermLevel = KEY_SET_VALUE
        Case 3 : lAccPermLevel = KEY_CREATE_SUB_KEY
        Case 4 : lAccPermLevel = DELETE
        End Select
        
        If Not RegCheckAccess(hDefKey, sSubKeyName, lAccPermLevel) Then
            fIsAdmin = False
            Exit for
        End If
    Next 'i
    
    If fIsAdmin Then
        sSubKeyName = "Software\Microsoft\Windows\"
        For i = 1 To 4
            Select Case i
            Case 1 : lAccPermLevel = KEY_QUERY_VALUE
            Case 2 : lAccPermLevel = KEY_SET_VALUE
            Case 3 : lAccPermLevel = KEY_CREATE_SUB_KEY
            Case 4 : lAccPermLevel = DELETE
            End Select
            
            If Not RegCheckAccess(hDefKey, sSubKeyName, lAccPermLevel) Then
                fIsElevated = False
                Exit for
            End If
        Next 'i
    End If 'fIsAdmin

    
    Set ShellApp = CreateObject("Shell.Application") 

    If Not sPreReqError = vbNullString Then
        If fQuiet = False Then
            Msgbox "Script execution needs to terminate" & vbCrLf & vbCrLf & sPreReqError, vbOkOnly, _
            "Critical Error in Script Prerequisite Check"
        End If
        CheckPreReq = False
    End If ' sPreReqError = vbNullString
End Function 'CheckPreReq()
'=======================================================================================================

Sub CheckPreReq_ErrorHandler

    sPreReqError = sPreReqError & _
        sDebugErr & " returned:" & vbCrLf &_
        "Error Details: " & Err.Source & " " & Hex( Err ) & ": " & Err.Description & vbCrLf & vbCrLf
End Sub

'=======================================================================================================
'Module ComputerProperties
'=======================================================================================================

Sub ComputerProperties
    Dim oOS, oOsItem
    Dim sOSinfo, sOSVersion, sUserInfo, sSubKeyName, sName, sValue, sOsMui, sOsLcid, sCulture
    Dim arrKeys, arrNames, arrTypes, arrVersion
    Dim qOS, OsLang, ValueType
    Dim iOSVersion, iValueName
    Dim hDefKey
    Const REG_SZ = 1
    On Error Resume Next
    
    sComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    
    'Note 64 bit OS check was already done in 'Sub Initialize'
    
    'OS info from WMI Win32_OperatingSystem
    'Set oWMI = GetObject("winmgmts:\\.\root\cimv2") 
    Set qOS = oWMILocal.ExecQuery("Select * from Win32_OperatingSystem") 
    For Each oOsItem in qOS 
        sOSinfo = sOSinfo & oOsItem.Caption 
        sOSinfo = sOSinfo & oOsItem.OtherTypeDescription
        sOSinfo = sOSinfo & CSV & "SP " & oOsItem.ServicePackMajorVersion
        sOSinfo = sOSinfo & CSV & "Version: " & oOsItem.Version
        sOsVersion = oOsItem.Version
        sOSinfo = sOSinfo & CSV & "Codepage: " & oOsItem.CodeSet
        sOSinfo = sOSinfo & CSV & "Country Code: " & oOsItem.CountryCode
        sOSinfo = sOSinfo & CSV & "Language: " & oOsItem.OSLanguage
    Next
    sOSinfo = sOSinfo & CSV & "System Type: " & sSystemType 
    
    'Build the VersionNT number
    arrVersion = Split(sOsVersion, Delimiter(sOsVersion))
    iVersionNt = CInt(arrVersion(0))*100 + CInt(arrVersion(1))
    
    hDefKey = HKEY_LOCAL_MACHINE 
    If iVersionNt < 600 Then 
        ' "old" reg location
        sSubKeyName = "System\CurrentControlSet\Control\Nls\MUILanguages\" 
        If RegEnumValues (hDefKey, sSubKeyName, arrNames, arrTypes) Then
            For iValueName = 0 To UBound(arrNames) 
                If arrTypes(iValueName) = REG_SZ Then
                    sName = "": sName = arrNames(iValueName)
                    sOsMui = sOsMui & GetCultureInfo(CInt("&h"& sName)) & " (" & CInt("&h"& sName) & ")" & CSV
                End If
            Next 'ValueName
        sOsMui = RTrimComma(sOsMui) 
        End If 'IsArray
    Else
        ' "new" reg location
        sSubKeyName = "SYSTEM\CurrentControlSet\Control\MUI\UILanguages\"
        sName = "LCID"
        If RegEnumKey(hDefKey, sSubKeyName, arrKeys) Then
            For Each OsLang in arrKeys 
                If Len(sOsMui) > 1 Then sOsMui = sOsMui & CSV
                sOsMui = sOsMui & OsLang 
                If RegReadDWordValue(hDefKey, sSubKeyName&OsLang, sName, sValue) Then sOsMui = sOsMui & " (" & sValue & ")"
            Next 'OsLang
        End If 'RegKeyExists
    End If
    
    'User info from WMI Win32_ComputerSystem
    Set qOS = oWMILocal.ExecQuery("Select * from Win32_ComputerSystem") 
    For Each oOsItem in qOS
        sUserInfo = sUserInfo & "Username: " & oOsItem.UserName 
    Next
    sUserInfo = sUserInfo & CSV & "IsAdmin: " & fIsAdmin
    sUserInfo = sUserInfo & CSV & "SID: " & sCurUserSid
    
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "Windows Installer Version", oMsi.Version
    CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "ComputerName", sComputerName
    CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "OS Details", sOSinfo 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "OS MUI Languages", sOsMui 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "Current User", sUserInfo 
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "Logfile Name", sLogFile
    CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "ROI Script Build", SCRIPTBUILD
End Sub 'ComputerProperties

'=======================================================================================================
'Module Logging
'=======================================================================================================

'=======================================================================================================

Sub WriteLog
    Dim LogOutput
    Dim i, n
    Dim sScriptSettings
    On Error Resume Next
    
    'Add actual values for customizable settings
    sScriptSettings = ""
    If fListNonOfficeProducts Then sScriptSettings = sScriptSettings & "/All "
    If fLogFull Then sScriptSettings = sScriptSettings & "/Full "
    If fLogVerbose Then sScriptSettings = sScriptSettings & "/LogVerbose "
    If fLogChainedDetails Then sScriptSettings = sScriptSettings & "/LogChainedDetails "
    If fFileInventory Then sScriptSettings = sScriptSettings & "/FileInventory "
    If fFeatureTree Then sScriptSettings = sScriptSettings & "/FeatureTree "
    If fBasicMode Then sScriptSettings = "/Basic "
    If fQuiet Then sScriptSettings = sScriptSettings & "/Quiet "
    If Not sPathOutputFolder = "" Then sScriptSettings = sScriptSettings & "/Logfolder: " & sPathOutputFolder
    CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "Script Settings", sScriptSettings 
    
    'Determine total scan time
    tEnd = Time()
    If NOT fBasicMode Then CacheLog LOGPOS_COMPUTER, LOGHEADING_NONE, "Total scan time", Int((tEnd - tStart) * 1000000 + 0.5) / 10 & " s"

    Set LogOutput = oFso.OpenTextFile(sLogFile, FOR_WRITING, True, True)
    LogOutput.WriteLine "Microsoft Customer Support Services - Robust Office Inventory - " & Now
    LogOutput.Write vbCrLf & String(160, "*") & vbCrLf

    'Flush the prepared output arrays to the log
    n = 2
    'Enable raw output if an error was encountered
    If NOT Len(arrLog(1)) = 32 AND NOT fBasicMode Then n = 3
    For i = 0 to n
        If i = 1 Then
            If NOT Len(arrLog(i)) = 32 Then 
                LogOutput.Write arrLog(i)
                LogOutput.Write vbCrLf & String(160, "*") & vbCrLf
            End If
        Else
            LogOutput.Write arrLog(i)
            If NOT (i = 2) Then LogOutput.Write vbCrLf & String(160, "*") & vbCrLf
        End If
    Next 'i
    LogOutput.Close
    Set LogOutput = Nothing
    
    'Copy the log to the fileinventory folder if needed
    If fFileInventory Then oFso.CopyFile sLogFile, sPathOutputFolder & "ROIScan\", TRUE

    ' write the xml log
    WriteXmlLog

End Sub
'=======================================================================================================

Sub WriteXmlLog
    Dim XmlLogStream, mspSeq, key, lic, component, item
    Dim sXmlLine, sText, sProductCode, sSPLevel, sFamily, sSeq, sItem, sTmp
    Dim i, iVPCnt, iItem, iArpCnt, iDummy, iPos, iPosMaster, iChainProd, iPosPatch, iColPatch
    Dim iColISource, iLogCnt
    Dim arrC2RPackages, arrC2RItems, arrTmp, arrTmpInner, arrLicData, arrKeyComponents
    Dim arrComponentData, arrVAppState
    Dim dicXmlTmp, dicApps
    Dim fLogProduct

    On Error Resume Next
    
    Set dicXmlTmp = CreateObject("Scripting.Dictionary")
    Set dicApps = CreateObject("Scripting.Dictionary")

    Set XmlLogStream = oFso.CreateTextFile(sPathOutputFolder & sComputerName & "_ROIScan.xml", True, True)
    XmlLogStream.WriteLine "<?xml version= ""1.0""?>"
    XmlLogStream.WriteLine "<OFFICEINVENTORY>"

    'c2r v3
    If UBound(arrVirt3Products) > -1 Then
        sXmlLine = "<C2RShared "
        For Each key in dicC2RPropV3.Keys
            sItem = dicC2RPropV3.Item(key)
            sItem = Replace(sItem, STR_NOTCONFIGURED, STR_NOTCONFIGUREDXML)
            sXmlLine = sXmlLine & " " & Replace(key, " ","") & "=" & chr(34) & sItem & chr(34)
        Next
        ' end line
        sXmlLine = sXmlLine & " />"
        'flush
        XmlLogStream.WriteLine sXmlLine

        For iVPCnt = 0 To UBound(arrVirt3Products, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirt3Products (iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' KeyName
            sXmlLine = sXmlLine & " KeyName=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_KEYNAME) & chr(34)
            ' ConfigName
            sXmlLine = sXmlLine & " ConfigName=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_CONFIGNAME) & chr(34)
            sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " ProductBuild=" & chr(34) & dicC2RPropV3.Item(STR_BUILDNUMBER) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " Version=" & chr(34) & arrVirt3Products (iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' Architecture
            sXmlLine = sXmlLine & " Architecture=" & chr(34) & dicC2RPropV3.Item(STR_PLATFORM) & chr(34)
            ' InstallType
            sXmlLine = sXmlLine & " InstallType=" & chr(34) & "C2R" & chr(34)
            ' C2R integration ProductCode
            sXmlLine = sXmlLine & " C2rIntegrationProductCode=" & chr(34) & dicC2RPropV3.Item(STR_PACKAGEGUID) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine

            ' Child Packages
            XmlLogStream.WriteLine "<ChildPackages>"
            arrC2RPackages = Split (arrVirt3Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
            If UBound(arrC2RPackages) > 0 Then
                For i = 0 To UBound(arrC2RPackages) - 1 'strip off last delimiter
                    arrC2RItems = Split(arrC2RPackages(i), " - ")
                    If UBound(arrC2RItems) = 1 Then
                        If NOT i = 0 Then XmlLogStream.WriteLine "</SkuCulture>"
                        XmlLogStream.WriteLine "<SkuCulture culture=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " >"
                    Else
                        XmlLogStream.WriteLine "<ChildPackage ProductCode=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " ProductName=" & chr(34) & arrC2RItems(2) & chr(34) & " />"
                    End If
                Next 'i
                XmlLogStream.WriteLine "</SkuCulture>"
            End If 'UBound(arrC2RPackages) > 0
            XmlLogStream.WriteLine "</ChildPackages>"
            
            'KeyComponents
            MapAppsToConfigProduct dicApps, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME), 16
            XmlLogStream.WriteLine "<KeyComponents>"
            For Each key in dicApps
                Set arrVAppState = Nothing
                sTmp = key & "_" & dicC2RPropV3.Item(STR_PLATFORM)
                If dicKeyComponentsV3.Exists(sTmp) Then
                    arrVAppState = Split(dicKeyComponentsV3.Item(sTmp), ",")
                    sXmlLine = "<Application "
                    sXmlLine = sXmlLine & " Name=" & chr(34) & key & chr(34)
                    sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrVAppState (1) & chr(34)
                    sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrVAppState (2) & chr(34)
                    sXmlLine = sXmlLine & " Version=" & chr(34) & arrVAppState (3) & chr(34)
                    sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrVAppState (4) & chr(34)
                    sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrVAppState (5) & chr(34)
                    sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrVAppState (6) & chr(34)
                    sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrVAppState (7) & chr(34)
                    sXmlLine = sXmlLine & " Path=" & chr(34) & arrVAppState (8) & chr(34)
                    sXmlLine = sXmlLine & " />"
                    XmlLogStream.WriteLine sXmlLine
                End If
            Next 'key
            XmlLogStream.WriteLine "</KeyComponents>"

            'LicenseData
            XmlLogStream.WriteLine "<LicenseData>"
            XmlLogStream.WriteLine arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSEXML)
            XmlLogStream.WriteLine "</LicenseData>"
            
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v2


    'c2r v2
    If UBound(arrVirt2Products) > -1 Then
        sXmlLine = "<C2RShared "
        For Each key in dicC2RPropV2.Keys
            sItem = dicC2RPropV2.Item(key)
            sItem = Replace(sItem, STR_NOTCONFIGURED, STR_NOTCONFIGUREDXML)
            sXmlLine = sXmlLine & " " & Replace(key, " ","") & "=" & chr(34) & sItem & chr(34)
        Next
        ' end line
        sXmlLine = sXmlLine & " />"
        'flush
        XmlLogStream.WriteLine sXmlLine

        For iVPCnt = 0 To UBound(arrVirt2Products, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirt2Products (iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' KeyName
            sXmlLine = sXmlLine & " KeyName=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_KEYNAME) & chr(34)
            ' ConfigName
            sXmlLine = sXmlLine & " ConfigName=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_CONFIGNAME) & chr(34)
            sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " BuildNumber=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_PRODUCTVERSION) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " Version=" & chr(34) & arrVirt2Products (iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' Architecture
            sXmlLine = sXmlLine & " Architecture=" & chr(34) & dicC2RPropV2.Item(STR_PLATFORM) & chr(34)
            ' InstallType
            sXmlLine = sXmlLine & " InstallType=" & chr(34) & "C2R" & chr(34)
            ' C2R integration ProductCode
            sXmlLine = sXmlLine & " C2rIntegrationProductCode=" & chr(34) & dicC2RPropV2.Item(STR_PACKAGEGUID) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine

            ' Child Packages
            XmlLogStream.WriteLine "<ChildPackages>"
            arrC2RPackages = Split (arrVirt2Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
            If UBound(arrC2RPackages) > 0 Then
                For i = 0 To UBound(arrC2RPackages) - 1 'strip off last delimiter
                    arrC2RItems = Split(arrC2RPackages(i), " - ")
                    If UBound(arrC2RItems) = 1 Then
                        If NOT i = 0 Then XmlLogStream.WriteLine "</SkuCulture>"
                        XmlLogStream.WriteLine "<SkuCulture culture=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " >"
                    Else
                        XmlLogStream.WriteLine "<ChildPackage ProductCode=" & chr(34) & arrC2RItems(0) & chr(34) & " ProductVersion=" & chr(34) & arrC2RItems(1) & chr(34) & " ProductName=" & chr(34) & arrC2RItems(2) & chr(34) & " />"
                    End If
                Next 'i
                XmlLogStream.WriteLine "</SkuCulture>"
            End If 'UBound(arrC2RPackages) > 0
            XmlLogStream.WriteLine "</ChildPackages>"
            
            'KeyComponents
            MapAppsToConfigProduct dicApps, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME), 15
            XmlLogStream.WriteLine "<KeyComponents>"
            For Each key in dicApps
                Set arrVAppState = Nothing
                sTmp = key & "_" & dicC2RPropV2.Item(STR_PLATFORM)
                If dicKeyComponentsV2.Exists(sTmp) Then
                    arrVAppState = Split(dicKeyComponentsV2.Item(sTmp), ",")
                    sXmlLine = "<Application "
                    sXmlLine = sXmlLine & " Name=" & chr(34) & key & chr(34)
                    sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrVAppState (1) & chr(34)
                    sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrVAppState (2) & chr(34)
                    sXmlLine = sXmlLine & " Version=" & chr(34) & arrVAppState (3) & chr(34)
                    sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrVAppState (4) & chr(34)
                    sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrVAppState (5) & chr(34)
                    sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrVAppState (6) & chr(34)
                    sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrVAppState (7) & chr(34)
                    sXmlLine = sXmlLine & " Path=" & chr(34) & arrVAppState (8) & chr(34)
                    sXmlLine = sXmlLine & " />"
                    XmlLogStream.WriteLine sXmlLine
                End If
            Next 'key
            XmlLogStream.WriteLine "</KeyComponents>"
            
            'LicenseData
            XmlLogStream.WriteLine "<LicenseData>"
            XmlLogStream.WriteLine arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSEXML)
            XmlLogStream.WriteLine "</LicenseData>"
            
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v2

    'c2r v1
    If UBound(arrVirtProducts) > -1 Then
        For iVPCnt = 0 To UBound(arrVirtProducts, 1)
            sXmlLine = ""
            sXmlLine = "<SKU "
            ' ProductName (heading)
            sXmlLine = sXmlLine & "ProductName=" & chr(34) & arrVirtProducts(iVPCnt, COL_PRODUCTNAME) & chr(34)
            ' ProductVersion
            sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrVirtProducts(iVPCnt, VIRTPROD_PRODUCTVERSION) & chr(34)
            ' SP Level
            sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrVirtProducts(iVPCnt, VIRTPROD_SPLEVEL) & chr(34)
            ' ProductCode
            sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrVirtProducts(iVPCnt, COL_PRODUCTCODE) & chr(34)
            ' end line
            sXmlLine = sXmlLine & " >"
            'flush
            XmlLogStream.WriteLine sXmlLine
            XmlLogStream.WriteLine "</SKU>"
        Next 'iVPCnt
    End If 'c2r v1

    'MSI Office > v12 Products
    For iArpCnt = 0 To UBound(arrArpProducts)
        iPos = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, ARP_CONFIGPRODUCTCODE))
        sProductCode = "" 
        If NOT iPos = -1 Then sProductCode = arrMaster(iPos, COL_PRODUCTCODE)
        sXmlLine = ""
        sXmlLine = "<SKU "
        ' ProductName (heading)
        sText = "" : sText = GetArpProductName(arrArpProducts(iArpCnt, COL_CONFIGNAME)) 
        sXmlLine = sXmlLine & "ProductName=" & chr(34) & sText & chr(34)
        ' Configuration ProductName
        sText = "" : sText = arrArpProducts(iArpCnt, COL_CONFIGNAME)
        If InStr(sText, ".") > 0 Then sText = Mid(sText, InStr(sText, ".") + 1)
        sXmlLine = sXmlLine & " ConfigName=" & chr(34) & sText & chr(34)
        sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
        ' ProductCode
        sXmlLine = sXmlLine & " ProductCode=" & chr(34) & "" & chr(34)
        ' ProductVersion
        sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrArpProducts(iArpCnt, ARP_PRODUCTVERSION) & chr(34)
        ' ServicePack
        sSPLevel = ""
        If NOT iPos = -1 AND NOT sProductCode = "" AND Len(arrArpProducts(iArpCnt, ARP_PRODUCTVERSION)) > 2 Then
            sSPLevel = OVersionToSpLevel(sProductCode, GetVersionMajor(sProductCode), arrArpProducts(iArpCnt, ARP_PRODUCTVERSION))
        End If
        sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & sSPLevel  & chr(34)
        ' Architecture
        sXmlLine = sXmlLine & " Architecture=" & chr(34) & arrMaster(iPos, COL_ARCHITECTURE) & chr(34)
        ' InstallType
        sXmlLine = sXmlLine & " InstallType=" & chr(34) & arrMaster(iPos, COL_INSTALLTYPE) & chr(34)
        ' InstallDate
        sXmlLine = sXmlLine & " InstallDate=" & chr(34) & Left(Replace(arrMaster(iPos, COL_INSTALLDATE), " ",""), 8) & chr(34)
        ' ProductState
        sXmlLine = sXmlLine & " ProductState=" & chr(34) & arrMaster(iPos, COL_STATESTRING) & chr(34)
        ' ConfigMsi
        sXmlLine = sXmlLine & " ConfigMsi=" & chr(34) & arrMaster(iPos, COL_ORIGINALMSI) & chr(34)
        ' BuildOrigin
        sXmlLine = sXmlLine & " BuildOrigin=" & chr(34) & arrMaster(iPos, COL_ORIGIN) & chr(34)
        ' end line
        sXmlLine = sXmlLine & " >"
        'flush
        XmlLogStream.WriteLine sXmlLine

        ' Child Packages
        XmlLogStream.WriteLine "<ChildPackages>"
            For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
                If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then
                    'log the master entry only to ensure a consistent xml data structure
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & sProductCode & chr(34)
                    'ProductVersion
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                    'ProductName
                    sText = "" : sText = arrMaster(iPosMaster, COL_ARPPRODUCTNAME)
                    If sText = "" Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                    Exit For
                Else
                    iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    'Only run if iPosMaster has a valid index #
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTCODE)
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & sText & chr(34)
                    'ProductVersion
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTVERSION)
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & sText & chr(34)
                    'ProductName
                    sText = "" : If NOT iPosMaster = -1 Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                End If
            Next 'iChainProd
        XmlLogStream.WriteLine "</ChildPackages>"

        'KeyComponents
        If Len(arrMaster(iPos, COL_KEYCOMPONENTS)) > 0 Then
            If Right(arrMaster(iPos, COL_KEYCOMPONENTS), 1) = ";" Then arrMaster(iPos, COL_KEYCOMPONENTS) = Left(arrMaster(iPos, COL_KEYCOMPONENTS), Len(arrMaster(iPos, COL_KEYCOMPONENTS)) - 1 )
        End If
        XmlLogStream.WriteLine "<KeyComponents>"
        arrKeyComponents = Split(arrMaster(iPos, COL_KEYCOMPONENTS), ";")
        For Each component in arrKeyComponents
            arrComponentData = Split(component, ",")
            If CheckArray(arrComponentData) Then
            sXmlLine = "<Application "
            sXmlLine = sXmlLine & " Name=" & chr(34) & arrComponentData(0) & chr(34)
            sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrComponentData(1) & chr(34)
            sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrComponentData(2) & chr(34)
            sXmlLine = sXmlLine & " Version=" & chr(34) & arrComponentData(3) & chr(34)
            sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrComponentData(4) & chr(34)
            sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrComponentData(5) & chr(34)
            sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrComponentData(6) & chr(34)
            sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrComponentData(7) & chr(34)
            sXmlLine = sXmlLine & " Path=" & chr(34) & arrComponentData(8) & chr(34)
            sXmlLine = sXmlLine & " />"
            XmlLogStream.WriteLine sXmlLine
            End If
        Next 'component
        XmlLogStream.WriteLine "</KeyComponents>"

        'LicenseData
        '-----------
        XmlLogStream.WriteLine "<LicenseData>"
        If NOT iPos = -1 Then
            If arrMaster(iPos, COL_OSPPLICENSE) <> "" Then
                XmlLogStream.WriteLine arrMaster(iPos, COL_OSPPLICENSEXML)
            End If
        End If
        XmlLogStream.WriteLine "</LicenseData>"

        ' Patches
        XmlLogStream.WriteLine "<PatchData>"
        ' PatchBaseLines
        XmlLogStream.WriteLine "<PatchBaseline Sequence=" & chr(34) & arrArpProducts(iArpCnt, ARP_PRODUCTVERSION) & chr(34) & " >"
        dicXmlTmp.RemoveAll
        For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
            If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
            iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
            'Only run if iPosMaster has a valid index #
            If Not iPosMaster = -1 Then
                Set arrTmp = Nothing
                arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                For Each MspSeq in arrTmp
                    arrTmpInner = Split(MspSeq, ":")
                    sFamily = "" : sSeq = ""
                    sFamily = arrTmpInner(0)
                    sSeq  = arrTmpInner(1)
                    If (sSeq>arrMaster(iPosMaster, COL_PRODUCTVERSION)) Then 
                        If dicXmlTmp.Exists(sFamily) Then
                            If (sSeq > dicXmlTmp.Item(sFamily)) Then dicXmlTmp.Item(sFamily) =sSeq
                        Else
                            dicXmlTmp.Add sFamily, sSeq
                        End If
                    End If
                Next 'MspSeq
            End If 'Not iPosMaster = -1
        Next 'iChainProd
        For Each key in dicXmlTmp.Keys
            XmlLogStream.WriteLine "<PostBaseline PatchFamily=" & chr(34) & key & chr(34) & " Sequence=" & chr(34) & dicXmlTmp.Item(key) & chr(34) & " />"
        Next 'key
        XmlLogStream.WriteLine "</PatchBaseline>"

        ' PatchList
        For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
            If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
            iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
            'Only run if iPosMaster has a valid index #
            If Not iPosMaster = -1 Then
                For iPosPatch = 0 to UBound(arrPatch, 3)
                    If Not IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                        sXmlLine = "<Patch PatchedProduct=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                        For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                            If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then
                                sXmlLine =  sXmlLine & " " & Replace(arrLogFormat(ARRAY_PATCH, iColPatch), ": ","") & "=" & chr(34) & arrPatch(iPosMaster, iColPatch, iPosPatch) & chr(34)
                            End If
                        Next 'iColPatch
                        sXmlLine =  sXmlLine & " />"
                        XmlLogStream.WriteLine sXmlLine
'                            Set arrTmp = Nothing
'                            arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
'                            For Each MspSeq in arrTmp
'                                arrTmpInner = Split(MspSeq, ":")
'                                XmlLogStream.WriteLine "<MsiPatchSequence PatchFamily=" & chr(34) & arrTmpInner(0) & chr(34) & " Sequence=" & chr(34) & arrTmpInner(1) & chr(34) & " />"
'                            Next 'MspSeq
'                            XmlLogStream.WriteLine "</Patch>"
                    End If 'IsEmpty
                Next 'iPosPatch
            End If ' Not iPosMaster = -1
        Next 'iChainProd

        XmlLogStream.WriteLine "</PatchData>"

        'InstallSource
        XmlLogStream.WriteLine "<InstallSource>"
        For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
            If NOT iPos = -1 Then
                If Not IsEmpty(arrIS(iPos, iColISource)) Then _ 
                XmlLogStream.WriteLine "<Source " & Replace(arrLogFormat(ARRAY_IS, iColISource), " ","") & "=" & chr(34) & arrIS(iPos, iColISource) & chr(34) & " />"
            End If
        Next 'iColISource
        XmlLogStream.WriteLine "</InstallSource>"

        XmlLogStream.WriteLine "</SKU>"
'        End If
    Next 'iArpCnt

    'Other Products
    Err.Clear
    For iLogCnt = 0 To 12
        For iPosMaster = 0 To UBound(arrMaster)
            fLogProduct = CheckLogProduct(iLogCnt, iPosMaster)
            If fLogProduct Then
                'arrMaster contents
                sXmlLine = ""
                sXmlLine = "<SKU "
                ' ProductName (heading)
                sText = "" : sText = arrMaster(iPosMaster, COL_ARPPRODUCTNAME)
                If sText = "" Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                sXmlLine = sXmlLine & "ProductName=" & chr(34) & sText & chr(34)
                sXmlLine = sXmlLine & " ConfigName=" & chr(34) & "" & chr(34)
                Select Case iLogCnt
                Case 1, 3, 5
                    sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "TRUE" & chr(34)
                Case Else
                    sXmlLine = sXmlLine & " IsChainedChild=" & chr(34) & "FALSE" & chr(34)
                End Select
                ' ProductCode
                sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                ' ProductVersion
                sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                ' ServicePack
                sXmlLine = sXmlLine & " ServicePackLevel=" & chr(34) & arrMaster(iPosMaster, COL_SPLEVEL) & chr(34)
                ' Architecture
                sXmlLine = sXmlLine & " Architecture=" & chr(34) & arrMaster(iPosMaster, COL_ARCHITECTURE) & chr(34)
                ' InstallType
                sXmlLine = sXmlLine & " InstallType=" & chr(34) & arrMaster(iPosMaster, COL_INSTALLTYPE) & chr(34)
                ' InstallDate
                sXmlLine = sXmlLine & " InstallDate=" & chr(34) & Left(Replace(arrMaster(iPosMaster, COL_INSTALLDATE), " ",""), 8) & chr(34)
                ' ProductState
                sXmlLine = sXmlLine & " ProductState=" & chr(34) & arrMaster(iPosMaster, COL_STATESTRING) & chr(34)
                ' ConfigMsi
                sXmlLine = sXmlLine & " ConfigMsi=" & chr(34) & arrMaster(iPosMaster, COL_ORIGINALMSI) & chr(34)
                ' BuildOrigin
                sXmlLine = sXmlLine & " BuildOrigin=" & chr(34) & arrMaster(iPosMaster, COL_ORIGIN) & chr(34)
                ' end line
                sXmlLine = sXmlLine & " >"
                'flush
                XmlLogStream.WriteLine sXmlLine

                ' Child Packages
                XmlLogStream.WriteLine "<ChildPackages>"
                    'log the master entry only to ensure a consistent xml data structure
                    sXmlLine =  "<ChildPackage"
                    'ProductCode
                    sXmlLine = sXmlLine & " ProductCode=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                    'ProductVersion
                    sXmlLine = sXmlLine & " ProductVersion=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34)
                    'ProductName
                    sText = "" : sText = arrMaster(iPosMaster, COL_ARPPRODUCTNAME)
                    If sText = "" Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                    sXmlLine = sXmlLine & " ProductName=" & chr(34) & sText & chr(34)
                    ' end line
                    sXmlLine = sXmlLine & " />"
                    'flush
                    XmlLogStream.WriteLine sXmlLine
                XmlLogStream.WriteLine "</ChildPackages>"

                'KeyComponents
                If Len(arrMaster(iPosMaster, COL_KEYCOMPONENTS)) > 0 Then
                    If Right(arrMaster(iPosMaster, COL_KEYCOMPONENTS), 1) = ";" Then arrMaster(iPosMaster, COL_KEYCOMPONENTS) = Left(arrMaster(iPosMaster, COL_KEYCOMPONENTS), Len(arrMaster(iPosMaster, COL_KEYCOMPONENTS)) - 1 )
                End If
                XmlLogStream.WriteLine "<KeyComponents>"
                arrKeyComponents = Split(arrMaster(iPosMaster, COL_KEYCOMPONENTS), ";")
                For Each component in arrKeyComponents
                    arrComponentData = Split(component, ",")
                    If CheckArray(arrComponentData) Then
                    sXmlLine = "<Application "
                    sXmlLine = sXmlLine & " Name=" & chr(34) & arrComponentData(0) & chr(34)
                    sXmlLine = sXmlLine & " ExeName=" & chr(34) & arrComponentData(1) & chr(34)
                    sXmlLine = sXmlLine & " VersionMajor=" & chr(34) & arrComponentData(2) & chr(34)
                    sXmlLine = sXmlLine & " Version=" & chr(34) & arrComponentData(3) & chr(34)
                    sXmlLine = sXmlLine & " InstallState=" & chr(34) & arrComponentData(4) & chr(34)
                    sXmlLine = sXmlLine & " InstallStateString=" & chr(34) & arrComponentData(5) & chr(34)
                    sXmlLine = sXmlLine & " ComponentId=" & chr(34) & arrComponentData(6) & chr(34)
                    sXmlLine = sXmlLine & " FeatureName=" & chr(34) & arrComponentData(7) & chr(34)
                    sXmlLine = sXmlLine & " Path=" & chr(34) & arrComponentData(8) & chr(34)
                    sXmlLine = sXmlLine & " />"
                    XmlLogStream.WriteLine sXmlLine
                    End If
                Next 'component
                XmlLogStream.WriteLine "</KeyComponents>"

                ' Patches
                XmlLogStream.WriteLine "<PatchData>"

                ' PatchBaseLines
                XmlLogStream.WriteLine "<PatchBaseline Sequence=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTVERSION) & chr(34) & " >"
                dicXmlTmp.RemoveAll
                Set arrTmp = Nothing
                arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                For Each MspSeq in arrTmp
                    arrTmpInner = Split(MspSeq, ":")
                    sFamily = "" : sSeq = ""
                    sFamily = arrTmpInner(0)
                    sSeq  = arrTmpInner(1)
                    If (sSeq>arrMaster(iPosMaster, COL_PRODUCTVERSION)) Then 
                        If dicXmlTmp.Exists(sFamily) Then
                            If (sSeq > dicXmlTmp.Item(sFamily)) Then dicXmlTmp.Item(sFamily) = sSeq
                        Else
                            dicXmlTmp.Add sFamily, sSeq
                        End If
                    End If
                Next 'MspSeq
                For Each key in dicXmlTmp.Keys
                    XmlLogStream.WriteLine "<PostBaseline PatchFamily=" & chr(34) & key & chr(34) & " Sequence=" & chr(34) & dicXmlTmp.Item(key) & chr(34) & " />"
                Next 'key
                XmlLogStream.WriteLine "</PatchBaseline>"

                For iPosPatch = 0 to UBound(arrPatch, 3)
                    If Not IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                        sXmlLine = "<Patch PatchedProduct=" & chr(34) & arrMaster(iPosMaster, COL_PRODUCTCODE) & chr(34)
                        For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                            If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then
                                sXmlLine =  sXmlLine & " " & Replace(arrLogFormat(ARRAY_PATCH, iColPatch), ": ","") & "=" & chr(34) & arrPatch(iPosMaster, iColPatch, iPosPatch) & chr(34)
                            End If
                        Next 'iColPatch
                        sXmlLine =  sXmlLine & " />"
                        XmlLogStream.WriteLine sXmlLine
                    End If 'IsEmpty
                Next 'iPosPatch
                
                XmlLogStream.WriteLine "</PatchData>"

                'InstallSource
                XmlLogStream.WriteLine "<InstallSource>"
                For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                    If Not IsEmpty(arrIS(iPos, iColISource)) Then _ 
                    XmlLogStream.WriteLine "<Source " & Replace(arrLogFormat(ARRAY_IS, iColISource), " ","") & "=" & chr(34) & arrIS(iPosMaster, iColISource) & chr(34) & " />"
                Next 'iColISource
                XmlLogStream.WriteLine "</InstallSource>"

                XmlLogStream.WriteLine "</SKU>"
            End If 'fLogProduct
        Next 'iPosMaster
    Next 'iLogCnt

    XmlLogStream.WriteLine "</OFFICEINVENTORY>"

End Sub 'WriteXmlLog

'=======================================================================================================

Sub PrepareLog (sLogFormat)
    Dim Key, MspSeq, Lic, ScenarioItem
    Dim i, j, k, m, iLBound, iUBound
    Dim iAip, iPos, iPosMaster, iPosArp, iArpCnt, iChainProd, iPosOrder, iDummy
    Dim iPosPatch, iColPatch, iColISource, iPosISource, iLogCnt, iVPCnt
    Dim sTmp, sText, sTextErr, sTextNotes, sCategory, sSeq, sFamily, sFamilyMain, sFamilyLang
    Dim sProdCache, sLcid, sVer
    Dim bCspCondition, fLogProduct, fDataIntegrity
    Dim fLoggedVirt, fLoggedMulti, fLoggedSingle
    Dim arrOrder(), arrTmp, arrTmpInner, arrLicData, arrC2RPackages, arrVAppState
    Dim dicFamilyLang, dicApps, dicTmp, dicMspFamVer
    On Error Resume Next

    fDataIntegrity = True
    fLoggedMulti = False
    fLoggedSingle = False
    fLoggedVirt = False

    'Filter contents for the log
    i = -1
    Redim arrOrder(19)
    i = i + 1 : arrOrder(i)  = COL_PRODUCTVERSION  'BuildNumber
    i = i + 1 : arrOrder(i)  = COL_SPLEVEL         'Version
    i = i + 1 : arrOrder(i)  = COL_ARCHITECTURE
    i = i + 1 : arrOrder(i)  = COL_INSTALLTYPE
    i = i + 1 : arrOrder(i)  = COL_PRODUCTCODE
    i = i + 1 : arrOrder(i)  = COL_PRIMARYPRODUCTID
    i = i + 1 : arrOrder(i)  = COL_PRODUCTNAME
    i = i + 1 : arrOrder(i)  = COL_ARPPARENTS
    i = i + 1 : arrOrder(i)  = COL_INSTALLDATE
    i = i + 1 : arrOrder(i)  = COL_USERSID
    i = i + 1 : arrOrder(i)  = COL_CONTEXTSTRING
    i = i + 1 : arrOrder(i)  = COL_STATESTRING
    i = i + 1 : arrOrder(i) = COL_TRANSFORMS
    i = i + 1 : arrOrder(i) = COL_ORIGINALMSI
    i = i + 1 : arrOrder(i) = COL_CACHEDMSI
    i = i + 1 : arrOrder(i) = COL_PRODUCTID
    i = i + 1 : arrOrder(i) = COL_ORIGIN
    i = i + 1 : arrOrder(i) = COL_PACKAGECODE
    i = i + 1 : arrOrder(i) = COL_NOTES
    i = i + 1 : arrOrder(i) = COL_ERROR

' trim content fields
    For iPosMaster = 0 To UBound(arrMaster)
        arrMaster(iPosMaster, COL_NOTES) = Trim(arrMaster(iPosMaster, COL_NOTES)) 
        arrMaster(iPosMaster, COL_ERROR) = Trim(arrMaster(iPosMaster, COL_ERROR))
        If Right(arrMaster(iPosMaster, COL_NOTES), 1) = "," Then arrMaster(iPosMaster, COL_NOTES) = Left(arrMaster(iPosMaster, COL_NOTES), Len(arrMaster(iPosMaster, COL_NOTES)) - 1)
        If Right(arrMaster(iPosMaster, COL_ERROR), 1) = "," Then arrMaster(iPosMaster, COL_ERROR) = Left(arrMaster(iPosMaster, COL_ERROR), Len(arrMaster(iPosMaster, COL_ERROR)) - 1)
        If arrMaster(iPosMaster, COL_NOTES) = "" Then arrMaster(iPosMaster, COL_NOTES) = "-"
        If arrMaster(iPosMaster, COL_ERROR) = "" Then arrMaster(iPosMaster, COL_ERROR) = "-"
        If Left(arrMaster(iPosMaster, COL_ARPPARENTS), 1) = "," Then arrMaster(iPosMaster, COL_ARPPARENTS) =Replace(arrMaster(iPosMaster, COL_ARPPARENTS), ",", "", 1, 1, 1)
    Next 'iPosMaster
    
    Set dicTmp = CreateObject("Scripting.Dictionary")
    Set dicFamilyLang = CreateObject("Scripting.Dictionary")
    Set dicMspFamVer = CreateObject("Scripting.Dictionary")
    Set dicApps = CreateObject("Scripting.Dictionary")
    
' prepare Virtualized C2R Products
' --------------------------------
    If UBound(arrVirtProducts) > -1 OR UBound(arrVirt2Products) > -1 OR UBound(arrVirt3Products) > -1 Then
        CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "C2R Products" 
        fLoggedVirt = True
        
        'O16 C2R
        If UBound(arrVirt3Products) > -1 Then
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "C2R Shared Properties" 
                
            'Shared Properties
            For Each key in dicC2RPropV3.Keys
                Select Case key
                Case STR_REGPACKAGEGUID
                Case STR_UPDATECHANNEL
                Case STR_POLUPDATECHANNEL
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, "CDN: " & dicC2RPropV3.Item(key)
                Case STR_CDNBASEURL
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Install",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case STR_UPDATESENABLED
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Updates",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case STR_UPDATETOVERSION
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                Case Else
                    sTmp = dicC2RPropV3.Item(key)
                    If NOT sTmp = "" Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV3.Item(key)
                End Select
            Next


            ' Scenario key state
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "Scenario Key State" 
            For Each ScenarioItem in dicScenarioV3.Keys
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, ScenarioItem, dicScenarioV3.Item(ScenarioItem)
            Next 'ScenarioItem

            sProdCache = ""
            For iVPCnt = 0 To UBound(arrVirt3Products, 1)
                If NOT InStr(sProdCache, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)) > 0 Then
                    sProdCache = sProdCache & arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME) & ","

                    ' ProductName (heading)
                    sText = "" : sText = arrVirt3Products(iVPCnt, COL_PRODUCTNAME)
                    If InStr(sText, " - ") > 0 AND (Len(sText) = InStr(sText, " - ") + 7) Then sText = Left(sText, InStr(sText, " - ") - 1)
                    ' Add C2R tag
                    sText = sText & " (C2R)"
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 
                    ' ProductVersion (BuildNumber)
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_PRODUCTVERSION)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION), sText
                    ' SP Level (Version)
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_SPLEVEL)
                    ' ConfigName
                    sText = "" : sText = arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME), sText
                    ' Languages
                    i = 0
                    For Each key in dicVirt3Cultures.Keys
                        If i = 0 Then sText = "Language(s)" Else sText = ""
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sText, key
                        i = i + 1
                    Next 'key
                    
                    'Application States
                    '------------------
                    MapAppsToConfigProduct dicApps, arrVirt3Products(iVPCnt, VIRTPROD_CONFIGNAME), 16
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    For Each key in dicApps
                        Set arrVAppState = Nothing
                        sTmp = key & "_" & dicC2RPropV3.Item(STR_PLATFORM)
                        If dicKeyComponentsV3.Exists(sTmp) Then
                            sText = "Installed - "
                            arrVAppState = Split(dicKeyComponentsV3.Item(sTmp), ",")
                            If NOT arrVAppState(4) = "3" Then
                                sText = "Absent/Excluded - "
                            Else
                                sText = sText & arrVAppState(3) & " - " & arrVAppState(6) & " - " & arrVAppState(7) & " - " & arrVAppState(8)
                            End If
                        End If
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, sText
                    Next 'key
                    
                    'LicenseData
                    '-----------
                    If arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSE) <> "" Then
                        arrLicData = Split(arrVirt3Products(iVPCnt, VIRTPROD_OSPPLICENSE), "#;#")
                        If CheckArray(arrLicData) Then
                            If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                            i = 0
                            For Each Lic in arrLicData
                                arrTmp = Split(Lic, ";")
                                If LCase(arrTmp(0)) = "active license" Then i = 1
                                If i < 2 Then
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrTmp(0), arrTmp(1)
                                Else
                                    If NOT (fBasicMode AND i > 5) Then
                                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrTmp(0) & ":" & Space(33 - Len(arrTmp(0))) & arrTmp(1)
                                    End If
                                End If
                                i = i + 1
                            Next 'Lic
                        End If 'arrLicData
                    End If 'arrMaster
                End If 'sProdCache
            Next 'iVPCnt
        'end C2R v3
        End If

        'O15 C2R
        If UBound(arrVirt2Products) > -1 Then
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "C2R Shared O15 Properties" 

            'Shared Properties
            For Each key in dicC2RPropV2.Keys
                Select Case key
                Case STR_REGPACKAGEGUID
                Case STR_CDNBASEURL
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Install",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case STR_UPDATESENABLED
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Updates",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case STR_UPDATETOVERSION
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                Case Else
                    sTmp = dicC2RPropV2.Item(key)
                    If NOT sTmp = "" Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, dicC2RPropV2.Item(key)
                End Select
            Next


            ' Scenario key state
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, "Scenario Key State" 
            For Each ScenarioItem in dicScenarioV2.Keys
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, ScenarioItem, dicScenarioV2.Item(ScenarioItem)
            Next 'ScenarioItem

            sProdCache = ""
            For iVPCnt = 0 To UBound(arrVirt2Products, 1)
                If NOT InStr(sProdCache, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)) > 0 Then
                    sProdCache = sProdCache & arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME) & ","

                    ' ProductName (heading)
                    sText = "" : sText = arrVirt2Products(iVPCnt, COL_PRODUCTNAME)
                    ' Add C2R tag
                    sText = sText & " (C2R)"
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 
                    ' ProductVersion
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_PRODUCTVERSION)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION), sText
                    ' SP Level
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_SPLEVEL)
                    ' ConfigName
                    sText = "" : sText = arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_CONFIGNAME), sText
                    ' Languages
                    i = 0
                    For Each key in dicVirt2Cultures.Keys
                        If i = 0 Then sText = "Language(s)" Else sText = ""
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sText, key
                        i = i + 1
                    Next 'key
                    ' Child Packages
                    arrC2RPackages = Split(arrVirt2Products(iVPCnt, VIRTPROD_CHILDPACKAGES), ";")
                    For i = 0 To UBound(arrC2RPackages)
                        If i = 0 Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "Chained Packages", arrC2RPackages(i) _
                        Else CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrC2RPackages(i)
                    Next 'i
                    'Application States
                    '------------------
                    MapAppsToConfigProduct dicApps, arrVirt2Products(iVPCnt, VIRTPROD_CONFIGNAME), 15
                    For Each key in dicApps
                        Set arrVAppState = Nothing
                        sTmp = key & "_" & dicC2RPropV2.Item(STR_PLATFORM)
                        If dicKeyComponentsV2.Exists(sTmp) Then
                            sText = "Installed - "
                            arrVAppState = Split(dicKeyComponentsV2.Item(sTmp), ",")
                            If NOT arrVAppState(4) = "3" Then
                                sText = "Absent/Excluded - "
                            Else
                                sText = sText & arrVAppState(3) & " - " & arrVAppState(6) & " - " & arrVAppState(7) & " - " & arrVAppState(8)
                            End If
                        End If
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, key, sText
                    Next 'key

                    'LicenseData
                    '-----------
                    If arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSE) <> "" Then
                        arrLicData = Split(arrVirt2Products(iVPCnt, VIRTPROD_OSPPLICENSE), "#;#")
                        If CheckArray(arrLicData) Then
                            If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                            i = 0
                            For Each Lic in arrLicData
                                arrTmp = Split(Lic, ";")
                                If LCase(arrTmp(0)) = "active license" Then i = 1
                                If i < 2 Then
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrTmp(0), arrTmp(1)
                                Else
                                    If NOT (fBasicMode AND i > 5) Then
                                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrTmp(0) & ":" & Space(33 - Len(arrTmp(0))) & arrTmp(1)
                                    End If
                                End If
                                i = i + 1
                            Next 'Lic
                        End If 'arrLicData
                    End If 'arrMaster
                End If 'sProdCache
            Next 'iVPCnt
        End If

        If UBound(arrVirtProducts) > -1 Then
            For iVPCnt = 0 To UBound(arrVirtProducts, 1)
                sText = "" : sText = arrVirtProducts(iVPCnt, COL_PRODUCTNAME)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 
                sText = "" : sText = arrVirtProducts(iVPCnt, VIRTPROD_PRODUCTVERSION)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_PRODUCTVERSION), sText
                sText = "" : sText = arrVirtProducts(iVPCnt, VIRTPROD_SPLEVEL)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, VIRTPROD_SPLEVEL), sText
                sText = "" : sText = arrVirtProducts(iVPCnt, COL_PRODUCTCODE)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_VIRTPROD, COL_PRODUCTCODE), sText
            Next 'iVPCnt
        End If
        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", vbCrLf & String(160, "*") & vbCrLf
    Else
        'No virtualized products found
    End If
    
    'Prepare Office > 12 Products
    '----------------------------
    If CheckArray(arrArpProducts) Then
        CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "Chained Products View" 
        fLoggedMulti = True

    For iArpCnt = 0 To UBound(arrArpProducts)
    ' extra loop to allow exit out of a bad product
        For iDummy = 1 To 1
        ' get the link to the configuration .MSI
            iPos = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, ARP_CONFIGPRODUCTCODE))
            If iPos = -1 Then
                If NOT dicMissingChild.Exists(arrArpProducts(iArpCnt, ARP_CONFIGPRODUCTCODE)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
            End If
            'If possible use the entry as shown under ARP as heading for the log
            sText = "" : sText = GetArpProductName(arrArpProducts(iArpCnt, COL_CONFIGNAME)) 
            If (sText = "" AND iPos = -1) Then sText = arrMaster(iPos, COL_ARPPRODUCTNAME)
            sText = sText & " (MSI)"
            CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText 

            'Contents from master array
            '--------------------------
            For iPosOrder = 0 To UBound(arrOrder) - 2 '-2 to exclude Notes and Error fields
                If fBasicMode AND iPosOrder > 1 Then Exit For
                sText = "" : If NOT iPos=- 1 Then sText = arrMaster(iPos, arrOrder(iPosOrder)) 
                'Suppress Configuration SKU for ARP config products. Log all others
                If NOT arrOrder(iPosOrder) = COL_ARPPARENTS Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, arrOrder(iPosOrder)), sText
            Next 'iPosOrder

            If NOT fBasicMode Then
                'Add Notes and Errors from chained products to ArpProd column
                sTextNotes = "" : sTextErr = ""
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
                    If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    If iPosMaster = -1 AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt, iChainProd)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        If Not arrMaster(iPosMaster, COL_NOTES) = "-" Then
                            sTextNotes = sTextNotes & arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster, COL_NOTES)
                        End If 'arrMaster(iPosMaster, COL_NOTES) = "-"
                        If Not arrMaster(iPosMaster, COL_ERROR) = "-" Then
                            sTextErr = sTextErr & arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster, COL_ERROR)
                        End If 'arrMaster(iPosMaster, COL_NOTES) = "-"
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
                If sTextNotes = "" Then sTextNotes = "-"
                If sTextErr = "" Then sTextErr = "-"
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_NOTES), sTextNotes
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_ERROR), sTextErr
            
                'Configuration ProductName
                '-------------------------
                sText = "" : sText = arrArpProducts(iArpCnt, COL_CONFIGNAME)
                If InStr(sText, ".") > 0 Then sText = Mid(sText, InStr(sText, ".") + 1)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_ARP, COL_CONFIGNAME), sText
            
                'Configuration PackageID
                sText = "" : sText = arrArpProducts(iArpCnt, COL_CONFIGPACKAGEID)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_ARP, COL_CONFIGPACKAGEID), sText
            
                'Chained packages
                '----------------
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                sCategory = "Chained Packages"
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
                    If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt, iChainProd)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        sText = "" : sText = arrMaster(iPosMaster, COL_PRODUCTCODE) & DSV & arrMaster(iPosMaster, COL_PRODUCTVERSION) & DSV & arrMaster(iPosMaster, COL_PRODUCTNAME)
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, sText
                        sCategory = ""
                    Else
                        sText = "" : sText = arrArpProducts(iArpCnt, iChainProd) & DSV & ERR_CATEGORYERROR & "Missing chained product!"
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, sText
                        sCategory = ""
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
            End If 'fBasicMode

            'LicenseData
            '-----------
            If NOT iPos = -1 Then
                If arrMaster(iPos, COL_OSPPLICENSE) <> "" Then
                    arrLicData = Split(arrMaster(iPos, COL_OSPPLICENSE), "#;#")
                    If CheckArray(arrLicData) Then
                        If NOT fBasicMode Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                        i = 0
                        For Each Lic in arrLicData
                            arrTmp = Split(Lic, ";")
                            If LCase(arrTmp(0)) = "active license" Then i = 1
                            If i < 2 Then
                                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrTmp(0), arrTmp(1)
                            Else
                                If NOT (fBasicMode AND i > 5) Then
                                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", arrTmp(0)& ":"& Space(25-Len(arrTmp(0)))&arrTmp(1)
                                End If
                            End If
                            i = i + 1
                        Next 'Lic
                    End If 'arrLicData
                End If 'arrMaster
            End If 'iPos - 1
            
            If NOT fBasicMode Then
                'Patches
                '-------
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                sCategory = "Patch Baseline"
                sText = "" : If NOT iPos = -1 Then sText = arrMaster(iPos, COL_PRODUCTVERSION)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, sText
            
                sCategory = "Post Baseline Sequences"
                dicTmp.RemoveAll
                dicMspFamVer.RemoveAll
                For iChainProd = COL_LBOUNDCHAINLIST To UBound (arrArpProducts, 2)
                    If IsEmpty (arrArpProducts(iArpCnt, iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition (arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists (arrArpProducts(iArpCnt, iChainProd)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
                    
                    If Not iPosMaster = -1 Then
                        Set arrTmp = Nothing
                        arrTmp = Split (arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                        For Each MspSeq in arrTmp
                            arrTmpInner = Split (MspSeq, ":")
                            sFamily = "" : sSeq = "" : sFamilyMain = "" : sFamilyLang = ""
                            sFamily = arrTmpInner(0)
                            sSeq  = arrTmpInner(1)
                            If (sSeq > arrMaster(iPosMaster, COL_PRODUCTVERSION)) Then 
                                If dicTmp.Exists(sFamily) Then
                                    If (sSeq > dicTmp.Item(sFamily)) Then dicTmp.Item(sFamily) = sSeq
                                Else
                                    dicTmp.Add sFamily, sSeq
                                End If
                            End If
                        Next 'MspSeq
                    End If 'Not iPosMaster = -1
                Next 'iChainProd
                
                For Each key in dicTmp.Keys
                    sTmp = ""
                    If InStr (key, "_") > 0 Then
                       If Len (key) = InStr (key, "_") + 4 Then
                            sTmp = Left (key, InStr (key, "_")) & dicTmp.Item (key)
                            sLcid = Right (key, 4)
                            If NOT dicMspFamVer.Exists (sTmp) Then
                                dicMspFamVer.Add sTmp, sLcid
                            Else
                                If NOT InStr (dicMspFamVer.Item (sTmp), sLcid) > 0 Then dicMspFamVer.Item (sTmp) = dicMspFamVer.Item (sTmp) & "," & sLcid
                            End If
                        End If
                    Else
                        If NOT dicMspFamVer.Exists (key) Then dicMspFamVer.Add key, dicTmp.Item (key)
                    End If
                Next

                For Each key in dicMspFamVer.Keys
                    If InStr (key, "_") > 0 Then
                        arrTmp = Split (key, "_")
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, arrTmp (1) & " - " & arrTmp (0) & " (" & dicMspFamVer.Item (key) & ")"
                    Else
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, dicTmp.Item(Key) & " - " & key
                    End If
                    sCategory = ""
                Next 'key
            
                sCategory = "Patchlist by product"
                For iChainProd = COL_LBOUNDCHAINLIST To UBound(arrArpProducts, 2)
                    If IsEmpty(arrArpProducts(iArpCnt, iChainProd)) Then Exit For
                    iPosMaster = GetArrayPosition(arrMaster, arrArpProducts(iArpCnt, iChainProd))
                    If iPosMaster = -1  AND NOT dicMissingChild.Exists(arrArpProducts(iArpCnt, iChainProd)) Then Cachelog LOGPOS_REVITEM, LOGHEADING_NONE, ERR_CATEGORYERROR, ERR_DATAINTEGRITY
                    'Only run if iPosMaster has a valid index #
                    If Not iPosMaster = -1 Then
                        For iPosPatch = 0 to UBound(arrPatch, 3)
                            If Not IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch)) Then
                                If iPosPatch = 0 Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                                If iPosPatch = 0 Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, arrMaster(iPosMaster, COL_PRODUCTNAME)& " - "&arrMaster(iPosMaster, COL_PRODUCTCODE)
                                sCategory = ""
                                sText = "  "
                                For iColPatch = PATCH_LOGSTART to PATCH_LOGCHAINEDMAX
                                    If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then sText = sText & arrLogFormat(ARRAY_PATCH, iColPatch) & arrPatch(iPosMaster, iColPatch, iPosPatch) & CSV
                                Next 'iColPatch
                                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, RTrimComma(sText)
                            End If 'IsEmpty
                        Next 'iPosPatch
                    End If ' Not iPosMaster = -1
                Next 'iChainProd
            
                'InstallSource
                '-------------
                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                    If NOT iPos = -1 Then
                        If Not IsEmpty(arrIS(iPos, iColISource)) Then _ 
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_IS, iColISource), arrIS(iPos, iColISource)
                    End If
                Next 'iColISource
            End If 'fBasicMode

        Next 'iDummy
    Next 'iArpCnt
    If fLoggedMulti Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", vbCrLf & String(160, "*") & vbCrLf
    End If 'CheckArray(arrArpProducts)

    'Prepare Other Products
    '======================
    Err.Clear
    For iLogCnt = 0 To 12
        For iPosMaster = 0 To UBound(arrMaster)
            fLogProduct = CheckLogProduct(iLogCnt, iPosMaster)
            If fLogProduct Then
                If NOT fLoggedSingle Then CacheLog LOGPOS_PRODUCT, LOGHEADING_H1, Null, "Single .msi Products View" 
                fLoggedSingle = True
                For iDummy = 1 To 1
                'arrMaster contents
                '------------------
                sText = "" : sText = arrMaster(iPosMaster, COL_ARPPRODUCTNAME)
                If sText = "" Then sText = arrMaster(iPosMaster, COL_PRODUCTNAME)
                CacheLog LOGPOS_PRODUCT, LOGHEADING_H2, Null, sText
                If arrMaster(iPosMaster, COL_STATE) = INSTALLSTATE_UNKNOWN Then
                    sText = arrMaster(iPosMaster, COL_PRODUCTCODE)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_PRODUCTCODE), sText
                    sText = arrMaster(iPosMaster, COL_USERSID)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_USERSID), sText
                    sText = arrMaster(iPosMaster, COL_CONTEXTSTRING)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_CONTEXTSTRING), sText
                    sText = arrMaster(iPosMaster, COL_STATESTRING)
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, COL_STATESTRING), sText
                    Exit For 'Dummy
                End If 'arrMaster(iPosMster, COL_STATE) = INSTALLSTATE_UNKNOWN
                For iPosOrder = 0 To UBound(arrOrder)
                    If fBasicMode AND iPosOrder > 1 Then Exit For
                    sText = "" : sText = arrMaster(iPosMaster, arrOrder(iPosOrder))
                    If Not IsEmpty(arrMaster(iPosMaster, arrOrder(iPosOrder))) Then 
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_MASTER, arrOrder(iPosOrder)), sText
                        If arrMaster(iPosMaster, arrOrder(iPosOrder)) = "Unknown" Then Exit For
                    End If
                Next 'iPosOrder
                
                If NOT fBasicMode Then
                    'Patches
                    '-------
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    'First loop will take care of client side patches
                    'Second loop will log patches in the InstallSource

                    sCategory = "Patch Baseline"
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, arrMaster(iPosMaster, COL_PRODUCTVERSION)
                
                    sCategory = "Post Baseline Sequences"
                    dicTmp.RemoveAll
                    Set arrTmp = Nothing
                    arrTmp = Split(arrMaster(iPosMaster, COL_PATCHFAMILY), ",")
                    For Each MspSeq in arrTmp
                        arrTmpInner = Split(MspSeq, ":")
                        dicTmp.Add arrTmpInner(0), arrTmpInner(1)
                    Next 'MspSeq
                    For Each Key in dicTmp.Keys
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, dicTmp.Item(Key) & " - " & Key
                        sCategory = ""
                    Next
                
                    sCategory = "Client side patches"  
                    bCspCondition = True
                    For iAip = 0 To 1
                        For iPosPatch = 0 to UBound(arrPatch, 3)
                            If Not (IsEmpty (arrPatch(iPosMaster, PATCH_PATCHCODE, iPosPatch))) AND (arrPatch(iPosMaster, PATCH_CSP, iPosPatch) = bCspCondition) Then
                                sText = ""
                                For iColPatch = PATCH_LOGSTART to PATCH_LOGMAX
                                    If Not IsEmpty(arrPatch(iPosMaster, iColPatch, iPosPatch)) Then sText = sText & arrLogFormat(ARRAY_PATCH, iColPatch) & arrPatch(iPosMaster, iColPatch, iPosPatch) & CSV
                                Next 'iColPatch
                                CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, sCategory, sText
                                sCategory = ""
                            End If 'IsEmpty
                        Next 'iPosPatch
                        sCategory = "Patches in InstallSource"
                        bCspCondition = False
                    Next 'iAip

                    'InstallSource
                    '-------------
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    For iColISource = IS_LOG_LBOUND To IS_LOG_UBOUND
                        If Not IsEmpty(arrIS(iPosMaster, iColISource)) Then _ 
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, arrLogFormat(ARRAY_IS, iColISource), arrIS(iPosMaster, iColISource)
                    Next 'iColISource
                
                    'FeatureStates
                    '-------------
                    CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "",""
                    If fFeatureTree Then
                        CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "FeatureStates", arrFeature(iPosMaster, FEATURE_TREE)
                    End If 'fFeatureTree
                End If 'fBasicMode
            Next 'iDummy
            End If 'fLogProduct
        Next 'iPosMaster
    Next 'iLogCnt
    If fLoggedSingle Then CacheLog LOGPOS_PRODUCT, LOGHEADING_NONE, "", vbCrLf & String(160, "*") & vbCrLf


End Sub 'PrepareLog
'=======================================================================================================

Sub Cachelog (iArrLogPosition, iHeading, sCategory, sText)
    On Error Resume Next
    
    If Not iHeading = 0 Then sText = FormatHeading(iHeading, sText) 
    If Not IsNull(sCategory) Then sCategory = FormatCategory(sCategory) 
    If iArrLogPosition = LOGPOS_REVITEM Then 
        If Not InStr(arrLog(iArrLogPosition), sText) > 0 Then _
         arrLog(iArrLogPosition) = arrLog(iArrLogPosition) & sCategory & sText &vbCrLf 
    Else
        arrLog(iArrLogPosition) = arrLog(iArrLogPosition) & sCategory & sText &vbCrLf 
    End If
End Sub
'=======================================================================================================

Function CheckLogProduct (iLogCnt, iPosMaster)
    Dim fLogProduct

    Select Case iLogCnt
            
    Case 0 'Add-Ins
        fLogProduct = False
        fLogProduct = fLogProduct OR (InStr(POWERPIVOT_2010, arrMaster(iPosMaster, COL_UPGRADECODE)) > 0 AND Len(arrMaster(iPosMaster, COL_UPGRADECODE)) = 38)
    Case 1 'Office 16
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 16
    Case 2 'Office 16 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 16
        If arrMaster(iPosMaster, COL_INSTALLTYPE) = "C2R" AND NOT fLogVerbose Then fLogProduct = False
    Case 3 'Office 15 
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 15
    Case 4 'Office 15 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 15
        If arrMaster(iPosMaster, COL_INSTALLTYPE) = "C2R" AND NOT fLogVerbose Then fLogProduct = False
    Case 5 'Office 14
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 14
    Case 6 'Office 14 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 14
    Case 7 'Office 12
        fLogProduct = fLogChainedDetails AND arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 12
    Case 8 'Office 12 Single Msi Products
        fLogProduct = NOT fLogChainedDetails AND _ 
                        arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND _ 
                        ( (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 1) OR (NOT arrMaster(iPosMaster, COL_SYSTEMCOMPONENT) = 0 AND arrMaster(iPosMaster, COL_ARPPARENTCOUNT) = 0) ) AND _ 
                        IsEmpty(arrMaster(iPosMaster, COL_ARPPARENTCOUNT)) AND _ 
                        GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 12
    Case 9 'Office 11
        fLogProduct = arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 11
    Case 10 'Office 10
        fLogProduct = arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 10
    Case 11 'Office  9
        fLogProduct = arrMaster(iPosMaster, COL_ISOFFICEPRODUCT) AND GetVersionMajor(arrMaster(iPosMaster, COL_PRODUCTCODE)) = 9
    Case 12 'Non Office Products
        fLogProduct = fListNonOfficeProducts AND NOT arrMaster(iPosMaster, COL_ISOFFICEPRODUCT)
    Case Else
    End Select

    CheckLogProduct = fLogProduct
End Function 'CheckLogProduct

'=======================================================================================================

Function FormatCategory(sCategory)
    Dim sTmp : Dim i
    On Error Resume Next

    Const iCATLEN = 33
    If Len(sCategory) > iCATLEN - 1 Then sTmp = sTmp & vbTab Else _
        sTmp = sTmp & Space(iCATLEN - Len(sCategory) - 1)

    FormatCategory = sCategory & sTmp 
End Function
'=======================================================================================================

Function FormatHeading(iHeading, sText)
    Dim sTmp, sStyle 
    Dim i
    On Error Resume Next

    Select Case iHeading
    Case LOGHEADING_H1: sStyle = "="
    Case LOGHEADING_H2: sStyle = "-"
    Case Else: sStyle = " "
    End Select
    
    sTmp = sTmp & String(Len(sText), sStyle)
    FormatHeading = vbCrLf & vbCrLf & sText & vbCrlf & sTmp 

End Function

'=======================================================================================================
'Module Global Helper Functions
'=======================================================================================================

'=======================================================================================================

Sub RelaunchElevated
    Dim Argument
    Dim sCmdLine
    Dim oShell, oWShell

    Set oShell = CreateObject("Shell.Application")
    Set oWShell = CreateObject("Wscript.Shell")

    sCmdLine = Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            If Argument = "UAC" Then Exit Sub
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    oShell.ShellExecute oWShell.ExpandEnvironmentStrings("%windir%") & "\system32\cscript.exe", sCmdLine & " UAC","","runas", 1
    Wscript.Quit
End Sub 'RelaunchElevated
'=======================================================================================================

Sub RelaunchAsCScript
    Dim Argument
    Dim sCmdLine
    
    sCmdLine = oShell.ExpandEnvironmentStrings("%windir%") & "\system32\cscript.exe " & Chr(34) & WScript.scriptFullName & Chr(34)
    If Wscript.Arguments.Count > 0 Then
        For Each Argument in Wscript.Arguments
            sCmdLine = sCmdLine  &  " " & chr(34) & Argument & chr(34)
        Next 'Argument
    End If
    oShell.Run sCmdLine, 1, False
    Wscript.Quit
End Sub 'RelaunchAsCScript
'=======================================================================================================

'Launch a Shell command, wait for the task to complete and return the result
Function Command(sCommand)
    Dim oExec
    Dim sCmdOut

    Set oExec = oShell.Exec(sCommand)
    sCmdOut = oExec.StdOut.ReadAll()
    Do While oExec.Status = 0
         WScript.Sleep 100
    Loop
    Command = oExec.ExitCode & " - " & sCmdOut
End Function
'=======================================================================================================

'-------------------------------------------------------------------------------
'   GetFullPathFromRelative
'
'   Expands a relative path syntax to the full path
'-------------------------------------------------------------------------------
Function GetFullPathFromRelative (sRelativePath)
    Dim sScriptDir

    sScriptDir = Left (wscript.ScriptFullName, InStrRev (wscript.ScriptFullName, "\"))
    ' ensure sRelativePath has no leading "\"
    If Left (sRelativePath, 1) = "\" Then sRelativePath = Mid (sRelativePath, 2)
    GetFullPathFromRelative = oFso.GetAbsolutePathName (sScriptDir & sRelativePath)

End Function 'GetFullPathFromRelative

'=======================================================================================================

Sub CopyToZip (ShellNameSpace, fsoFile)
    Dim fCopyComplete
    Dim item, ShellFolder
    Dim i

    Set ShellFolder = ShellNameSpace

    fCopyComplete = False
    ShellNameSpace.CopyHere fsoFile.Path, COPY_OVERWRITE
    For Each item in ShellFolder.Items
        If item.Name = fsoFile.Name Then fCopyComplete = True
    Next 'item
    i = 0
    While NOT fCopyComplete
        WScript.Sleep 500
        i = i + 1
        For Each item in ShellFolder.Items
            If item.Name = fsoFile.Name Then fCopyComplete = True
        Next 'item
    ' hang protection
        If i > 12 Then
            fCopyComplete = True
            fZipError = True
        End If
    Wend
End Sub

'-------------------------------------------------------------------------------
'   AddToDic
'
'   Adds a key, item pair to a dictionary object
'-------------------------------------------------------------------------------
Sub AddToDic(oDic, sKey, sItem)
    If NOT oDic.Exists(sKey) Then oDic.Add sKey, sItem
End Sub 'AddToDic

'-------------------------------------------------------------------------------
'   AddOrUpdateDic
'
'   Adds a key, item pair to a dictionary object if it does not exist
'   Updates the item value if it already exists
'-------------------------------------------------------------------------------
Sub AddOrUpdateDic(oDic, sKey, sItem)
    If NOT oDic.Exists(sKey) Then
        oDic.Add sKey, sItem
    Else
        oDic.Item(sKey) = sItem
    End If
End Sub 'AddOrUpdateDic

'=======================================================================================================

'Function to compare to numbers of unspecified format
Function CompareVersion (sFile1, sFile2, bAllowBlanks)
'Return values:
'Left file version is lower than right file version     - 1
'Left file version is identical to right file version    0
'Left file version is higher than right file version     1
'Invalid comparison                                      2

    Dim file1, file2
    Dim sDelimiter
    Dim iCnt, iAsc, iMax, iF1, iF2
    Dim bLEmpty, bREmpty

    CompareVersion = 0
    bLEmpty = False
    bREmpty = False
    
    'Ensure valid inputs values
    On Error Resume Next
    If IsEmpty(sFile1) Then bLEmpty = True
    If IsEmpty(sFile2) Then bREmpty = True
    If sFile1 = "" Then bLEmpty = True
    If sFile2 = "" Then bREmpty = True
' don't allow alpha characters
    If Not bLEmpty Then
        For iCnt = 1 To Len(sFile1)
            iAsc = Asc(UCase(Mid(sFile1, iCnt, 1)))
            If (iAsc > 64) AND (iAsc < 91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    If Not bREmpty Then
        For iCnt = 1 To Len(sFile2)
            iAsc = Asc(UCase(Mid(sFile2, iCnt, 1)))
            If (iAsc > 64) AND (iAsc < 91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    
    If bLEmpty AND (NOT bREmpty) Then
        If bAllowBlanks Then CompareVersion = -1 Else CompareVersion = 2
        Exit Function
    End If
    
    If (NOT bLEmpty) AND bREmpty Then
        If bAllowBlanks Then CompareVersion = 1 Else CompareVersion = 2
        Exit Function
    End If
    
    If bLEmpty AND bREmpty Then
        CompareVersion = 2
        Exit Function
    End If
    
' if Files are identical we're already done
    If sFile1 = sFile2 Then Exit Function
' split the VersionString
    file1 = Split(sFile1, Delimiter(sFile1))
    file2 = Split(sFile2, Delimiter(sFile2))
' ensure we get the lower count
    iMax = UBound(file1)
    CompareVersion = -1
    If iMax > UBound(file2) Then 
        iMax = UBound(file2)
        CompareVersion = 1
    End If
' compare the file versions
    For iCnt = 0 To iMax
        iF1 = CLng(file1(iCnt))
        iF2 = CLng(file2(iCnt))
        If iF1 > iF2 Then
            CompareVersion = 1
            Exit For
        ElseIf iF1 < iF2 Then
            CompareVersion = -1
            Exit For
        End If
    Next 'iCnt
End Function

'=======================================================================================================

Function hAtS(arrHex)
    Dim c, s

    On Error Resume Next
    hAtS = "" : s = ""
    If NOT IsArray(arrHex) Then Exit Function
    For Each c in arrHex
        s = s & Chr(CInt("&h" & c))
    Next 'c
    hAts = s
End Function
'=======================================================================================================

Function Delimiter (sVersion)
    Dim iCnt, iAsc

    Delimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            Delimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function
'=======================================================================================================

'Get the culture info tag from LCID
Function GetCultureInfo (sLcid)
    Dim sLang

    Select Case UCase(Hex(CInt(sLcid)))
        Case "7F" : sLang = ""        'Invariant culture
        Case "36" : sLang = "af"	 ' Afrikaans
        Case "436" : sLang = "af-ZA"	 ' Afrikaans (South Africa)
        Case "1C" : sLang = "sq"	 ' Albanian
        Case "41C" : sLang = "sq-AL"	 ' Albanian (Albania)
        Case "1" : sLang = "ar"	 ' Arabic
        Case "1401" : sLang = "ar-DZ"	 ' Arabic (Algeria)
        Case "3C01" : sLang = "ar-BH"	 ' Arabic (Bahrain)
        Case "C01" : sLang = "ar-EG"	 ' Arabic (Egypt)
        Case "801" : sLang = "ar-IQ"	 ' Arabic (Iraq)
        Case "2C01" : sLang = "ar-JO"	 ' Arabic (Jordan)
        Case "3401" : sLang = "ar-KW"	 ' Arabic (Kuwait)
        Case "3001" : sLang = "ar-LB"	 ' Arabic (Lebanon)
        Case "1001" : sLang = "ar-LY"	 ' Arabic (Libya)
        Case "1801" : sLang = "ar-MA"	 ' Arabic (Morocco)
        Case "2001" : sLang = "ar-OM"	 ' Arabic (Oman)
        Case "4001" : sLang = "ar-QA"	 ' Arabic (Qatar)
        Case "401" : sLang = "ar-SA"	 ' Arabic (Saudi Arabia)
        Case "2801" : sLang = "ar-SY"	 ' Arabic (Syria)
        Case "1C01" : sLang = "ar-TN"	 ' Arabic (Tunisia)
        Case "3801" : sLang = "ar-AE"	 ' Arabic (U.A.E.)
        Case "2401" : sLang = "ar-YE"	 ' Arabic (Yemen)
        Case "2B" : sLang = "hy"	 ' Armenian
        Case "42B" : sLang = "hy-AM"	 ' Armenian (Armenia)
        Case "2C" : sLang = "az"	 ' Azeri
        Case "82C" : sLang = "az-Cyrl-AZ"	 ' Azeri (Azerbaijan, Cyrillic)
        Case "42C" : sLang = "az-Latn-AZ"	 ' Azeri (Azerbaijan, Latin)
        Case "2D" : sLang = "eu"	 ' Basque
        Case "42D" : sLang = "eu-ES"	 ' Basque (Basque)
        Case "23" : sLang = "be"	 ' Belarusian
        Case "423" : sLang = "be-BY"	 ' Belarusian (Belarus)
        Case "2" : sLang = "bg"	 ' Bulgarian
        Case "402" : sLang = "bg-BG"	 ' Bulgarian (Bulgaria)
        Case "3" : sLang = "ca"	 ' Catalan
        Case "403" : sLang = "ca-ES"	 ' Catalan (Catalan)
        Case "C04" : sLang = "zh-HK"	 ' Chinese (Hong Kong SAR, PRC)
        Case "1404" : sLang = "zh-MO"	 ' Chinese (Macao SAR)
        Case "804" : sLang = "zh-CN"	 ' Chinese (PRC)
        Case "4" : sLang = "zh-Hans"	 ' Chinese (Simplified)
        Case "1004" : sLang = "zh-SG"	 ' Chinese (Singapore)
        Case "404" : sLang = "zh-TW"	 ' Chinese (Taiwan)
        Case "7C04" : sLang = "zh-Hant"	 ' Chinese (Traditional)
        Case "1A" : sLang = "hr"	 ' Croatian
        Case "41A" : sLang = "hr-HR"	 ' Croatian (Croatia)
        Case "5" : sLang = "cs"	 ' Czech
        Case "405" : sLang = "cs-CZ"	 ' Czech (Czech Republic)
        Case "6" : sLang = "da"	 ' Danish
        Case "406" : sLang = "da-DK"	 ' Danish (Denmark)
        Case "65" : sLang = "dv"	 ' Divehi
        Case "465" : sLang = "dv-MV"	 ' Divehi (Maldives)
        Case "13" : sLang = "nl"	 ' Dutch
        Case "813" : sLang = "nl-BE"	 ' Dutch (Belgium)
        Case "413" : sLang = "nl-NL"	 ' Dutch (Netherlands)
        Case "9" : sLang = "en"	 ' English
        Case "C09" : sLang = "en-AU"	 ' English (Australia)
        Case "2809" : sLang = "en-BZ"	 ' English (Belize)
        Case "1009" : sLang = "en-CA"	 ' English (Canada)
        Case "2409" : sLang = "en-029"	 ' English (Caribbean)
        Case "1809" : sLang = "en-IE"	 ' English (Ireland)
        Case "2009" : sLang = "en-JM"	 ' English (Jamaica)
        Case "1409" : sLang = "en-NZ"	 ' English (New Zealand)
        Case "3409" : sLang = "en-PH"	 ' English (Philippines)
        Case "1C09" : sLang = "en-ZA"	 ' English (South Africa
        Case "2C09" : sLang = "en-TT"	 ' English (Trinidad and Tobago)
        Case "809" : sLang = "en-GB"	 ' English (United Kingdom)
        Case "409" : sLang = "en-US"	 ' English (United States)
        Case "3009" : sLang = "en-ZW"	 ' English (Zimbabwe)
        Case "25" : sLang = "et"	 ' Estonian
        Case "425" : sLang = "et-EE"	 ' Estonian (Estonia)
        Case "38" : sLang = "fo"	 ' Faroese
        Case "438" : sLang = "fo-FO"	 ' Faroese (Faroe Islands)
        Case "29" : sLang = "fa"	 ' Farsi
        Case "429" : sLang = "fa-IR"	 ' Farsi (Iran)
        Case "B" : sLang = "fi"	 ' Finnish
        Case "40B" : sLang = "fi-FI"	 ' Finnish (Finland)
        Case "C" : sLang = "fr"	 ' French
        Case "80C" : sLang = "fr-BE"	 ' French (Belgium)
        Case "C0C" : sLang = "fr-CA"	 ' French (Canada)
        Case "40C" : sLang = "fr-FR"	 ' French (France)
        Case "140C" : sLang = "fr-LU"	 ' French (Luxembourg)
        Case "180C" : sLang = "fr-MC"	 ' French (Monaco)
        Case "100C" : sLang = "fr-CH"	 ' French (Switzerland)
        Case "56" : sLang = "gl"	 ' Galician
        Case "456" : sLang = "gl-ES"	 ' Galician (Spain)
        Case "37" : sLang = "ka"	 ' Georgian
        Case "437" : sLang = "ka-GE"	 ' Georgian (Georgia)
        Case "7" : sLang = "de"	 ' German
        Case "C07" : sLang = "de-AT"	 ' German (Austria)
        Case "407" : sLang = "de-DE"	 ' German (Germany)
        Case "1407" : sLang = "de-LI"	 ' German (Liechtenstein)
        Case "1007" : sLang = "de-LU"	 ' German (Luxembourg)
        Case "807" : sLang = "de-CH"	 ' German (Switzerland)
        Case "8" : sLang = "el"	 ' Greek
        Case "408" : sLang = "el-GR"	 ' Greek (Greece)
        Case "47" : sLang = "gu"	 ' Gujarati
        Case "447" : sLang = "gu-IN"	 ' Gujarati (India)
        Case "D" : sLang = "he"	 ' Hebrew
        Case "40D" : sLang = "he-IL"	 ' Hebrew (Israel)
        Case "39" : sLang = "hi"	 ' Hindi
        Case "439" : sLang = "hi-IN"	 ' Hindi (India)
        Case "E" : sLang = "hu"	 ' Hungarian
        Case "40E" : sLang = "hu-HU"	 ' Hungarian (Hungary)
        Case "F" : sLang = "is"	 ' Icelandic
        Case "40F" : sLang = "is-IS"	 ' Icelandic (Iceland)
        Case "21" : sLang = "id"	 ' Indonesian
        Case "421" : sLang = "id-ID"	 ' Indonesian (Indonesia)
        Case "10" : sLang = "it"	 ' Italian
        Case "410" : sLang = "it-IT"	 ' Italian (Italy)
        Case "810" : sLang = "it-CH"	 ' Italian (Switzerland)
        Case "11" : sLang = "ja"	 ' Japanese
        Case "411" : sLang = "ja-JP"	 ' Japanese (Japan)
        Case "4B" : sLang = "kn"	 ' Kannada
        Case "44B" : sLang = "kn-IN"	 ' Kannada (India)
        Case "3F" : sLang = "kk"	 ' Kazakh
        Case "43F" : sLang = "kk-KZ"	 ' Kazakh (Kazakhstan)
        Case "57" : sLang = "kok"	 ' Konkani
        Case "457" : sLang = "kok-IN"	 ' Konkani (India)
        Case "12" : sLang = "ko"	 ' Korean
        Case "412" : sLang = "ko-KR"	 ' Korean (Korea)
        Case "40" : sLang = "ky"	 ' Kyrgyz
        Case "440" : sLang = "ky-KG"	 ' Kyrgyz (Kyrgyzstan)
        Case "26" : sLang = "lv"	 ' Latvian
        Case "426" : sLang = "lv-LV"	 ' Latvian (Latvia)
        Case "27" : sLang = "lt"	 ' Lithuanian
        Case "427" : sLang = "lt-LT"	 ' Lithuanian (Lithuania)
        Case "2F" : sLang = "mk"	 ' Macedonian
        Case "42F" : sLang = "mk-MK"	 ' Macedonian (Macedonia, FYROM)
        Case "3E" : sLang = "ms"	 ' Malay
        Case "83E" : sLang = "ms-BN"	 ' Malay (Brunei Darussalam)
        Case "43E" : sLang = "ms-MY"	 ' Malay (Malaysia)
        Case "4E" : sLang = "mr"	 ' Marathi
        Case "44E" : sLang = "mr-IN"	 ' Marathi (India)
        Case "50" : sLang = "mn"	 ' Mongolian
        Case "450" : sLang = "mn-MN"	 ' Mongolian (Mongolia)
        Case "14" : sLang = "no"	 ' Norwegian
        Case "414" : sLang = "nb-NO"	 ' Norwegian (Bokmål, Norway)
        Case "814" : sLang = "nn-NO"	 ' Norwegian (Nynorsk, Norway)
        Case "15" : sLang = "pl"	 ' Polish
        Case "415" : sLang = "pl-PL"	 ' Polish (Poland)
        Case "16" : sLang = "pt"	 ' Portuguese
        Case "416" : sLang = "pt-BR"	 ' Portuguese (Brazil)
        Case "816" : sLang = "pt-PT"	 ' Portuguese (Portugal)
        Case "46" : sLang = "pa"	 ' Punjabi
        Case "446" : sLang = "pa-IN"	 ' Punjabi (India)
        Case "18" : sLang = "ro"	 ' Romanian
        Case "418" : sLang = "ro-RO"	 ' Romanian (Romania)
        Case "19" : sLang = "ru"	 ' Russian
        Case "419" : sLang = "ru-RU"	 ' Russian (Russia)
        Case "4F" : sLang = "sa"	 ' Sanskrit
        Case "44F" : sLang = "sa-IN"	 ' Sanskrit (India)
        Case "C1A" : sLang = "sr-Cyrl-CS"	 ' Serbian (Serbia, Cyrillic)
        Case "81A" : sLang = "sr-Latn-CS"	 ' Serbian (Serbia, Latin)
        Case "1B" : sLang = "sk"	 ' Slovak
        Case "41B" : sLang = "sk-SK"	 ' Slovak (Slovakia)
        Case "24" : sLang = "sl"	 ' Slovenian
        Case "424" : sLang = "sl-SI"	 ' Slovenian (Slovenia)
        Case "A" : sLang = "es"	 ' Spanish
        Case "2C0A" : sLang = "es-AR"	 ' Spanish (Argentina)
        Case "400A" : sLang = "es-BO"	 ' Spanish (Bolivia)
        Case "340A" : sLang = "es-CL"	 ' Spanish (Chile)
        Case "240A" : sLang = "es-CO"	 ' Spanish (Colombia)
        Case "140A" : sLang = "es-CR"	 ' Spanish (Costa Rica)
        Case "1C0A" : sLang = "es-DO"	 ' Spanish (Dominican Republic)
        Case "300A" : sLang = "es-EC"	 ' Spanish (Ecuador)
        Case "440A" : sLang = "es-SV"	 ' Spanish (El Salvador)
        Case "100A" : sLang = "es-GT"	 ' Spanish (Guatemala)
        Case "480A" : sLang = "es-HN"	 ' Spanish (Honduras)
        Case "80A" : sLang = "es-MX"	 ' Spanish (Mexico)
        Case "4C0A" : sLang = "es-NI"	 ' Spanish (Nicaragua)
        Case "180A" : sLang = "es-PA"	 ' Spanish (Panama)
        Case "3C0A" : sLang = "es-PY"	 ' Spanish (Paraguay)
        Case "280A" : sLang = "es-PE"	 ' Spanish (Peru)
        Case "500A" : sLang = "es-PR"	 ' Spanish (Puerto Rico)
        Case "C0A" : sLang = "es-ES"	 ' Spanish (Spain)
        Case "380A" : sLang = "es-UY"	 ' Spanish (Uruguay)
        Case "200A" : sLang = "es-VE"	 ' Spanish (Venezuela)
        Case "41" : sLang = "sw"	 ' Swahili
        Case "441" : sLang = "sw-KE"	 ' Swahili (Kenya)
        Case "1D" : sLang = "sv"	 ' Swedish
        Case "81D" : sLang = "sv-FI"	 ' Swedish (Finland)
        Case "41D" : sLang = "sv-SE"	 ' Swedish (Sweden)
        Case "5A" : sLang = "syr"	 ' Syriac
        Case "45A" : sLang = "syr-SY"	 ' Syriac (Syria)
        Case "49" : sLang = "ta"	 ' Tamil
        Case "449" : sLang = "ta-IN"	 ' Tamil (India)
        Case "44" : sLang = "tt"	 ' Tatar
        Case "444" : sLang = "tt-RU"	 ' Tatar (Russia)
        Case "4A" : sLang = "te"	 ' Telugu
        Case "44A" : sLang = "te-IN"	 ' Telugu (India)
        Case "1E" : sLang = "th"	 ' Thai
        Case "41E" : sLang = "th-TH"	 ' Thai (Thailand)
        Case "1F" : sLang = "tr"	 ' Turkish
        Case "41F" : sLang = "tr-TR"	 ' Turkish (Turkey)
        Case "22" : sLang = "uk"	 ' Ukrainian
        Case "422" : sLang = "uk-UA"	 ' Ukrainian (Ukraine)
        Case "20" : sLang = "ur"	 ' Urdu
        Case "420" : sLang = "ur-PK"	 ' Urdu (Pakistan)
        Case "43" : sLang = "uz"	 ' Uzbek
        Case "843" : sLang = "uz-Cyrl-UZ"	 ' Uzbek (Uzbekistan, Cyrillic)
        Case "443" : sLang = "uz-Latn-UZ"	 ' Uzbek (Uzbekistan, Latin)
        Case "2A" : sLang = "vi"	 ' Vietnamese
        Case "42A" : sLang = "vi-VN"	 ' Vietnamese (Vietnam)
    Case Else : sLang = ""
    End Select
    GetCultureInfo = sLang
End Function
'=======================================================================================================

'Trim away trailing comma from string
Function RTrimComma (sString)
    sString = RTrim (sString)
    If Right(sString, 1) = "," Then sString = Left(sString, Len(sString) - 1)
    RTrimComma = sString
End Function
'=======================================================================================================

'Return the primary keys of a table by using the PrimaryKeys property of the database object
'in SQL ready syntax 
Function GetPrimaryTableKeys(MsiDb, sTable)
    Dim iKeyCnt
    Dim sPrimaryTmp
    Dim PrimaryKeys
    On Error Resume Next

    sPrimaryTmp = ""
    Set PrimaryKeys = MsiDb.PrimaryKeys(sTable)
    For iKeyCnt = 1 To PrimaryKeys.FieldCount
        sPrimaryTmp = sPrimaryTmp & "`"&PrimaryKeys.StringData(iKeyCnt)& "`, "
    Next 'iKeyCnt
    GetPrimaryTableKeys = Left(sPrimaryTmp, Len(sPrimaryTmp) -2)
End Function 'GetPrimaryTableKeys
'=======================================================================================================

'Return the Column schema definition of a table in SQL ready syntax
Function GetTableColumnDef(MsiDb, sTable)
    On Error Resume Next
    Dim sQuery, sColDefTmp
    Dim View, ColumnNames, ColumnTypes
    Dim iColCnt
    
    'Get the ColumnInfo details
    sColDefTmp = ""
    sQuery = "SELECT * FROM " & sTable
    Set View = MsiDb.OpenView(sQuery)
    View.Execute
    Set ColumnNames = View.ColumnInfo(MSICOLUMNINFONAMES)
    Set ColumnTypes = View.ColumnInfo(MSICOLUMNINFOTYPES)
    For iColCnt = 1 To ColumnNames.FieldCount
        sColDefTmp = sColDefTmp & ColDefToSql(ColumnNames.StringData(iColCnt), ColumnTypes.StringData(iColCnt)) & ","
    Next 'iColCnt
    View.Close
    
    GetTableColumnDef = Left(sColDefTmp, Len(sColDefTmp) -2)
    
End Function 'GetTableColumnDef
'=======================================================================================================

'Translate the column definition fields into SQL syntax
Function ColDefToSql(sColName, sColType)
    On Error Resume Next
    
    Dim iLen
    Dim sRight, sLeft, sSqlTmp

    iLen = Len(sColType)
    sRight = Right(sColType, iLen- 1)
    sLeft = Left(sColType, 1)
    sSqlTmp = "`"& sColName& "`"
    Select Case sLeft
    Case "s","S"
        's? String, variable length (?= 1-255) -> CHAR(#) or CHARACTER(#)
        's0 String, variable length -> LONGCHAR
        If sRight= "0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR("& sRight& ")"
        If sLeft = "s" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "l","L"
        'CHAR(#) LOCALIZABLE or CHARACTER(#) LOCALIZABLE
        If sRight= "0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR("& sRight& ")"
        If sLeft = "l" Then sSqlTmp = sSqlTmp & " NOT NULL"
        If sRight= "0" Then sSqlTmp = sSqlTmp & "  LOCALIZABLE" Else sSqlTmp = sSqlTmp & " LOCALIZABLE"
    Case "i","I"
        'i2 Short integer 
        'i4 Long integer 
        If sRight= "2" Then sSqlTmp = sSqlTmp & " SHORT" Else sSqlTmp = sSqlTmp & " LONG"
        If sLeft = "i" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "v","V"
        'v0 Binary Stream 
        sSqlTmp = sSqlTmp & " OBJECT"
        If sLeft = "v" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "g","G"
        'g? Temporary string (?= 0-255)
    Case "j","J"
        'j? Temporary integer (?= 0, 1, 2, 4)) 
    Case "o","O"
        'O0 Temporary object 
    Case Else
    End Select

    ColDefToSql = sSqlTmp

End Function 'ColDefToSql
'=======================================================================================================

'Initialize the dicKeyComponents dictionary to link the key components to the ComponentClients
Sub InitKeyComponents
    Dim CompId, CompClients, client
    Dim arrKeyComponents

    On Error Resume Next
    arrKeyComponents = Array (CID_ACC16_64, CID_ACC16_32, CID_ACC15_64, CID_ACC15_32, CID_ACC14_64, CID_ACC14_32, CID_ACC12, CID_ACC11, CID_XL16_64, CID_XL16_32, CID_XL15_64, CID_XL15_32, CID_XL14_64, CID_XL14_32, CID_XL12, CID_XL11, CID_GRV16_64, CID_GRV16_32, CID_GRV15_64, CID_GRV15_32, CID_GRV14_64, CID_GRV14_32, CID_GRV12, CID_LYN16_64, CID_LYN16_32, CID_LYN15_64, CID_LYN15_32, CID_ONE16_64, CID_ONE16_32, CID_ONE15_64, CID_ONE15_32, CID_ONE14_64, CID_ONE14_32, CID_ONE12, CID_ONE11, CID_MSO16_64, CID_MSO16_32, CID_MSO15_64, CID_MSO15_32, CID_MSO14_64, CID_MSO14_32, CID_MSO12, CID_MSO11, CID_OL16_64, CID_OL16_32, CID_OL15_64, CID_OL15_32, CID_OL14_64, CID_OL14_32, CID_OL12, CID_OL11, CID_PPT16_64, CID_PPT16_32, CID_PPT15_64, CID_PPT15_32, CID_PPT14_64, CID_PPT14_32, CID_PPT12, CID_PPT11, CID_PRJ16_64, CID_PRJ16_32, CID_PRJ15_64, CID_PRJ15_32, CID_PRJ14_64, CID_PRJ14_32, CID_PRJ12, CID_PRJ11, CID_PUB16_64, CID_PUB16_32, CID_PUB15_64, CID_PUB15_32, CID_PUB14_64, CID_PUB14_32, CID_PUB12, CID_PUB11, CID_IP16_64, CID_IP15_64, CID_IP15_32, CID_IP14_64, CID_IP14_32, CID_IP12, CID_IP11, CID_VIS16_64, CID_VIS16_32, CID_VIS15_64, CID_VIS15_32, CID_VIS14_64, CID_VIS14_32, CID_VIS12, CID_VIS11, CID_WD16_64, CID_WD16_32, CID_WD15_64, CID_WD15_32, CID_WD14_64, CID_WD14_32, CID_WD12, CID_WD11, CID_SPD16_64, CID_SPD16_32, CID_SPD15_64, CID_SPD15_32, CID_SPD14_64, CID_SPD14_32, CID_SPD12, CID_SPD11)
    On Error Resume Next
    For Each CompId in arrKeyComponents
        Set CompClients = oMsi.ComponentClients (CompId)
        For Each client in CompClients
            If NOT CompClients.Count > 0 Then Exit For
            If NOT client = "" Then
                If NOT dicKeyComponents.Exists (CompId) Then
                    dicKeyComponents.Add CompId, client
                Else
                    dicKeyComponents.Item (CompId) = dicKeyComponents.Item (CompId) & ";" & client
                End If
            End If
        Next
    Next
End Sub 'InitKeyComponents

'=======================================================================================================

'Checks if a productcode belongs to a Office family
Function IsOfficeProduct (sProductCode)
    On Error Resume Next
    
    IsOfficeProduct = False
    If InStr(OFFICE_ALL, UCase(Right(sProductCode, 28))) > 0 OR _
       InStr(sProductCodes_C2R, UCase(sProductCode)) > 0 OR _
       InStr(OFFICEID, UCase(Right(sProductCode, 17))) > 0 OR _
       sProductCode = sPackageGuid Then
           IsOfficeProduct = True
    End If

End Function
'=======================================================================================================

Function GetExpandedGuid (sGuid)
    Dim sExpandGuid
    Dim i
    On Error Resume Next

    sExpandGuid = "{" & StrReverse(Mid(sGuid, 1, 8)) & "-" & _
                        StrReverse(Mid(sGuid, 9, 4)) & "-" & _
                        StrReverse(Mid(sGuid, 13, 4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid, (i + 1), 1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid, (i - 1), 1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid, (i + 1), 1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid, (i - 1), 1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "}"
    GetExpandedGuid = sExpandGuid
    
End Function
'=======================================================================================================

Function GetCompressedGuid (sGuid)
'Converts the GUID / ProductCode into the compressed format
    Dim sCompGUID
    Dim i
    On Error Resume Next

    sCompGUID = StrReverse(Mid(sGuid, 2, 8))  & _
                StrReverse(Mid(sGuid, 11, 4)) & _
                StrReverse(Mid(sGuid, 16, 4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function
'=======================================================================================================

'Get Version Major from GUID
Function GetVersionMajor(sProductCode)
    Dim iVersionMajor
    On Error Resume Next

    iVersionMajor = 0
    If InStr(OFFICE_2000, UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 9
    If InStr(ORK_2000,    UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 9
    If InStr(PRJ_2000,    UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 9
    If InStr(VIS_2002,    UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 10
    If InStr(OFFICE_2002, UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 10
    If InStr(OFFICE_2003, UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 11
    If InStr(WSS_2,       UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 11
    If InStr(SPS_2003,    UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 11
    If InStr(PPS_2007,    UCase(Right(sProductCode, 28))) > 0 Then iVersionMajor = 12
    If InStr(OFFICEID,    UCase(Right(sProductCode, 17))) > 0 Then iVersionMajor = Mid(sProductCode, 4, 2)
    
    If iVersionMajor = 0 Then iVersionMajor = oMsi.ProductInfo(sProductCode, "VersionMajor")  

    GetVersionMajor = iVersionMajor
End Function
'=======================================================================================================

'Obtain the ProductVersion from a .msi package
Function GetMsiProductVersion(sMsiFile)
    Dim MsiDb, Record
    Dim qView
    
    On Error Resume Next
    GetMsiProductVersion = ""
    Set Record = Nothing
    Set MsiDb = oMsi.OpenDatabase(sMsiFile, MSIOPENDATABASEMODE_READONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductVersion'")
    qView.Execute
    Set Record = qView.Fetch
    If NOT Record Is Nothing Then GetMsiProductVersion = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductVersion
'=======================================================================================================

'Get the reference name that is used to reference the product under add / remove programs
'Alternatively this could be taken from the cached .msi by reading the 'SetupExeArpId' value from the 'Property' table
Function GetArpProductname (sUninstallSubkey)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sVer
    On Error Resume Next
    
    GetArpProductname = sUninstallSubkey
    hDefKey = HKEY_LOCAL_MACHINE
    sSubKeyName = REG_ARP & sUninstallSubkey & "\"
    sName = "DisplayName"
    
    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then 
        If NOT IsNull(sValue) OR sValue = "" Then GetArpProductname = sValue
    Else
        'Try C2Rv2 location
        For Each sVer in dicActiveC2RVersions.Keys
            sSubKeyName = REG_OFFICE & sVer & REG_C2RVIRT_HKLM & sSubKeyName
            If RegReadValue (hDefKey, sSubKeyName, sName, sValue, "REG_EXPAND_SZ") Then
                If NOT IsNull(sValue) OR sValue = "" Then GetArpProductname = sValue
                Exit For
            End If
        Next
    End If

End Function
'=======================================================================================================

'Get the original .msi name (Package Name) by direct read from registry
Function GetRegOriginalMsiName(sProductCodeCompressed, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sFallBackName
    On Error Resume Next

    'PackageName is only available for per-machine, current user and managed user - not for other (unmanaged) user!
    
    sFallBackName = ""
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, False) & "SourceList\"
    Else
        'PackageName not available for other (unmanaged) user
        GetRegProductName = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid
        sName = "PackageName"

    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegOriginalMsiName = sValue  & sFallBackName Else GetRegOriginalMsiName = "-"

End Function 'GetRegOriginalMsiName
'=======================================================================================================

'Get the registered transform(s) by direct read from registry metadata
Function GetRegTransforms(sProductCodeCompressed, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue 
    On Error Resume Next

    'Transforms is only available for per-machine and current user - not for other user!
        
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, False)
        sName = "Transforms"
    Else
        'Transforms not available for other user
        GetRegTransforms = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid

    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegTransforms = sValue Else GetRegTransforms = "-"

End Function 'GetRegTransforms
'=======================================================================================================

'Get the product PackageCode by direct read from registry metadata
Function GetRegPackageCode(sProductCodeCompressed, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue 
    On Error Resume Next

    'PackageCode is only available for per-machine and current user - not for other user!
        
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        hDefKey = GetRegHive(iContext, sSid, False)
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, False)
        sName = "PackageCode"
    Else
        'PackageCode not available for other user
        GetRegPackageCode = "n/a"
        Exit Function
    End If 'sSid = sCurUserSid

    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegPackageCode = sValue Else GetRegPackageCode = "-"

End Function 'GetRegPackageCode
'=======================================================================================================

'Get the ProductName by direct read from registry metadata
Function GetRegProductName(sProductCodeCompressed, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue, sFallBackName
    Dim i
    On Error Resume Next

    'ProductName is only available for per-machine, current user and managed user - not for other (unmanaged) user!
    'If not per-machine, managed or SID not sCurUserSid tweak to 'DisplayName' from GlobalConfig key
        
    sFallBackName = ""
    If sSid = sCurUserSid Or iContext = MSIINSTALLCONTEXT_USERMANAGED Or iContext = MSIINSTALLCONTEXT_MACHINE Or iContext = MSIINSTALLCONTEXT_C2RV2 Or iContext = MSIINSTALLCONTEXT_C2RV3 Then
        If Not iContext = MSIINSTALLCONTEXT_USERMANAGED Then
            hDefKey = GetRegHive(iContext, sSid, False)
            sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, False)
        Else
            hDefKey = GetRegHive(iContext, sSid, True)
            sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, True)
        End If
        sName = "ProductName"
    Else
        'Use GlobalConfig key to avoid conflict with per-user installs
        hDefKey = HKEY_LOCAL_MACHINE
        sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iContext, sSid, True) & "InstallProperties\"
        sName = "DisplayName"
        sFallBackName = " (DisplayName)"
    End If 'sSid = sCurUserSid

    
    If RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue) Then GetRegProductName = sValue  & sFallBackName Else GetRegProductName = ERR_CATEGORYERROR & "No DislpayName registered"

End Function
'=======================================================================================================

Function GetRegProductState(sProductCodeCompressed, iContext, sSid)
    Dim hDefKey
    Dim sSubKeyName, sName, sValue
    Dim iTmpContext
    On Error Resume Next
    
    GetRegProductState = "Unknown"
    If iContext = MSIINSTALLCONTEXT_USERMANAGED Then 
        iTmpContext = MSIINSTALLCONTEXT_USERUNMANAGED
        sName = "ManagedLocalPackage"
    Else
        iTmpContext = iContext
        sName = "LocalPackage"
    End If 'iContext = MSIINSTALLCONTEXT_USERMANAGED
    sSubKeyName = GetRegConfigKey(sProductCodeCompressed, iTmpContext, sSid, True) & "InstallProperties\"
    
    If RegKeyExists (hDefKey, sSubKeyName) Then
        If RegValExists(hDefKey, sSubKeyName, sName) Then 
            GetRegProductState = 5 '"Installed"
        Else
            If InStr(sSubKeyName, REG_C2RVIRT_HKLM) > 0 Then
                GetRegProductState = 8 '"Virtualized"
            Else
                GetRegProductState = 1 '"Advertised"
            End If
        End If 'RegValExists
    Else
        GetRegProductState = -1 '"Broken/Unknown"
    End If 'RegKeyExists
End Function 'GetRegProductState
'=======================================================================================================

'Check if the GUID has a valid structure
Function IsValidGuid(sGuid, iGuidType)
    Dim i, n
    Dim c
    Dim bValidGuidLength, bValidGuidChar, bValidGuidCase
    On Error Resume Next

    'Set defaults
    IsValidGuid = False
    fGuidCaseWarningOnly = False
    sError = "" : sErrBpa = ""
    bValidGuidLength = True
    bValidGuidChar = True
   
    Select Case iGuidType
    Case GUID_UNCOMPRESSED 'UnCompressed
        If Len(sGuid) = 38 Then
            IsValidGuid = True
            For i = 1 To 38
                bValidGuidCase = True
                c = Mid(sGuid, i, 1)
                Select Case i
                Case 1
                    If Not c = "{" Then 
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case 10, 15, 20, 25
                    If not c = "-" Then
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case 38
                    If not c = "}" Then
                        IsValidGuid = False
                        bValidGuidChar = False
                        Exit For
                    End If
                Case Else
                    n = Asc(c)
                    If Not IsValidGuidChar(n) Then
                        bValidGuidCase = False
                        fGuidCaseWarningOnly = True
                        IsValidGuid = False
                        n = Asc(UCase(c))
                        If Not IsValidGuidChar(n) Then
                            fGuidCaseWarningOnly = False
                            bValidGuidChar = False
                            Exit For
                        End If 'IsValidGuidChar (inner)
                    End If 'IsValidGuidChar (outer)
                End Select
            Next 'i
        Else
            'Invalid length for this GUID type
            If NOT (Len(sGuid) =32) Then bValidGuidLength = False
        End If 'Len(sGuid)
        
    Case  GUID_COMPRESSED
        If Len(sGuid) =32 Then
            IsValidGuid = True
            For i = 1 to 32
                c = Mid(sGuid, i, 1)
                bValidGuidCase = True
                n = Asc(c)
                If Not IsValidGuidChar(n) Then 
                    bValidGuidCase = False
                    fGuidCaseWarningOnly = True
                    IsValidGuid = False
                    n = Asc(UCase(c))
                    If Not IsValidGuidChar(n) Then
                        fGuidCaseWarningOnly = False
                        bValidGuidChar = False
                         Exit For
                   End If 'IsValidGuidChar (inner)
                End If 'IsValidGuidChar (outer)
            Next 'i
        Else
            'Invalid length for this GUID type
            bValidGuidLength = False
        End If 'Len
    Case GUID_SQUISHED '"Squished"
        'Not implemented
    Case Else
        'IsValidGuid = False
    End Select
    
    'Log errors 
    If (NOT bValidGuidLength) OR (NOT bValidGuidChar) OR (fGuidCaseWarningOnly) Then 
         sError = ERR_INVALIDGUID & DOT
         sErrBpa = BPA_GUID
    End If 
    
    If fGuidCaseWarningOnly Then
        sError = sError & ERR_GUIDCASE
    End If 'fGuidCaseWarningOnly
    
    If bValidGuidLength = False Then 
        sError = sError & ERR_INVALIDGUIDLENGTH
    End If 'bValidGuidLength  
    
    If bValidGuidChar = False Then 
        sError = sError & ERR_INVALIDGUIDCHAR
    End If 'bValidGuidChar  
End Function 'IsValidGuid
'=======================================================================================================

'Check if the character is in a valid range for a GUID
Function IsValidGuidChar (iAsc)
    If ((iAsc >= 48 AND iAsc <= 57) OR (iAsc >= 65 AND iAsc <= 70)) Then
        IsValidGuidChar = True
    Else
        IsValidGuidChar = False
    End If
End Function
'=======================================================================================================

Function GetContextString(iContext)
    On Error Resume Next

    Select Case iContext
        Case MSIINSTALLCONTEXT_USERMANAGED      : GetContextString = "MSIINSTALLCONTEXT_USERMANAGED"
        Case MSIINSTALLCONTEXT_USERUNMANAGED    : GetContextString = "MSIINSTALLCONTEXT_USERUNMANAGED"
        Case MSIINSTALLCONTEXT_MACHINE          : GetContextString = "MSIINSTALLCONTEXT_MACHINE"
        Case MSIINSTALLCONTEXT_ALL              : GetContextString = "MSIINSTALLCONTEXT_ALL"
        Case MSIINSTALLCONTEXT_C2RV2            : GetContextString = "MSIINSTALLCONTEXT_C2RV2"
        Case MSIINSTALLCONTEXT_C2RV3            : GetContextString = "MSIINSTALLCONTEXT_C2RV3"
        Case Else                               : GetContextString = iContext
    End Select
End Function
'=======================================================================================================

Function GetHiveString(hDefKey)
    On Error Resume Next

    Select Case hDefKey
        Case HKEY_CLASSES_ROOT : GetHiveString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER : GetHiveString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE : GetHiveString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS : GetHiveString = "HKEY_USERS"
        Case Else : GetHiveString = hDefKey
    End Select
End Function 'GetHiveString
'=======================================================================================================

Function GetRegConfigPatchesKey(iContext, sSid, bGlobal)
    Dim sTmpProductCode, sSubKeyName, sKey, sVer
    On Error Resume Next

    sSubKeyName = ""
    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED
        sSubKeyName = REG_CONTEXTUSERMANAGED & sSid & "\Installer\Patches\"
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then
            sSubKeyName = REG_GLOBALCONFIG & sSid & "\Patches\"
        Else
            sSubKeyName = REG_CONTEXTUSER
        End If
    Case MSIINSTALLCONTEXT_MACHINE, MSIINSTALLCONTEXT_C2RV2, MSIINSTALLCONTEXT_C2RV3
        If bGlobal Then
            sSubKeyName = REG_GLOBALCONFIG & "S-1-5-18\Patches\"
        Else
            sSubKeyName = REG_CONTEXTMACHINE & "\Patches\"
        End If
    Case Else
    End Select
    
    GetRegConfigPatchesKey = Replace(sSubKeyName, "\\","\")
    
End Function 'GetRegConfigPatchesKey
'=======================================================================================================

Function GetRegConfigKey(sProductCode, iContext, sSid, bGlobal)
    Dim sTmpProductCode, sSubKeyName, sKey
    Dim iVM
    On Error Resume Next

    sTmpProductCode = sProductCode
    sSubKeyName = ""
    If NOT sTmpProductCode = "" Then
        If IsValidGuid(sTmpProductCode, GUID_UNCOMPRESSED) Then sTmpProductCode = GetCompressedGuid(sTmpProductCode)
    End If 'NOT sTmpProductCode = ""

    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED
        sSubKeyName = REG_CONTEXTUSERMANAGED & sSid & "\Installer\Products\" & sTmpProductCode & "\"
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then
            sSubKeyName = REG_GLOBALCONFIG & sSid & "\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = REG_CONTEXTUSER & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_MACHINE
        If bGlobal Then
            sSubKeyName = REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_C2RV2
        sKey = REG_OFFICE & "15.0" & REG_C2RVIRT_HKLM
        If bGlobal Then
            sSubKeyName = sKey & REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = sKey & "Software\Classes\" & REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case MSIINSTALLCONTEXT_C2RV3
        sKey = REG_OFFICE & REG_C2RVIRT_HKLM
        sKey = Replace(sKey, "\\","\")
        If bGlobal Then
            sSubKeyName = sKey & REG_GLOBALCONFIG & "S-1-5-18\Products\" & sTmpProductCode & "\"
        Else
            sSubKeyName = sKey & "Software\Classes\" & REG_CONTEXTMACHINE & "Products\" & sTmpProductCode & "\"
        End If
    Case Else
    End Select
    
    GetRegConfigKey = Replace(sSubKeyName, "\\","\")
    
End Function 'GetRegConfigKey
'=======================================================================================================

Function GetRegHive(iContext, sSid, bGlobal)
    On Error Resume Next

    
    Select Case iContext
    Case MSIINSTALLCONTEXT_USERMANAGED 
        GetRegHive = HKEY_LOCAL_MACHINE
    Case MSIINSTALLCONTEXT_USERUNMANAGED
        If bGlobal OR NOT sSid = sCurUserSid Then 
            GetRegHive = HKEY_LOCAL_MACHINE
        Else
            GetRegHive = HKEY_CURRENT_USER
        End If
    Case MSIINSTALLCONTEXT_MACHINE
        If bGlobal Then 
            GetRegHive = HKEY_LOCAL_MACHINE
        Else
            GetRegHive = HKEY_CLASSES_ROOT
        End If
    Case MSIINSTALLCONTEXT_C2RV2, MSIINSTALLCONTEXT_C2RV3
        GetRegHive = HKEY_LOCAL_MACHINE
    Case Else
    End Select
End Function 'GetRegHive
'=======================================================================================================

Function RegKeyExists(hDefKey, sSubKeyName)
    Dim arrKeys
    RegKeyExists = (oReg.EnumKey(hDefKey, sSubKeyName, arrKeys) = 0)
End Function
'=======================================================================================================

Function RegValExists(hDefKey, sSubKeyName, sName)
    Dim arrValueTypes, arrValueNames, i
    On Error Resume Next

    RegValExists = False
    If Not RegKeyExists(hDefKey, sSubKeyName) Then
        Exit Function
    End If
    If oReg.EnumValues(hDefKey, sSubKeyName, arrValueNames, arrValueTypes) = 0 AND CheckArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    Else
        Exit Function
    End If 'oReg.EnumValues
End Function
'=======================================================================================================

'Check access to a registry key
Function RegCheckAccess(hDefKey, sSubKeyName, lAccPermLevel)
    Dim RetVal
    Dim arrValues
    
    RetVal = RegKeyExists(hDefKey, sSubKeyName)
    RetVal = oReg.CheckAccess(hDefKey, sSubKeyName, lAccPermLevel)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.CheckAccess(hDefKey, Wow64Key(hDefKey, sSubKeyName), lAccPermLevel)
    RegCheckAccess = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

'Read the value of a given registry entry
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim arrValues
    
    Select Case UCase(sType)
        Case "1","REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "2","REG_EXPAND_SZ"
            RetVal = oReg.GetExpandedStringValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "7","REG_MULTI_SZ"
            RetVal = oReg.GetMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, arrValues)
            If RetVal = 0 Then sValue = Join(arrValues, chr(34))
        Case "4","REG_DWORD"
            RetVal = oReg.GetDWORDValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then 
                RetVal = oReg.GetDWORDValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
            End If
        Case "3","REG_BINARY"
            RetVal = oReg.GetBinaryValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case "11","REG_QWORD"
            RetVal = oReg.GetQWORDValue(hDefKey, sSubKeyName, sName, sValue)
            If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
        Case Else
            RetVal = -1
    End Select 'sValue
    RegReadValue = (RetVal = 0)
End Function 'RegReadValue
'=======================================================================================================

Function RegReadStringValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal

    RetVal = oReg.GetStringValue(hDefKey, sSubKeyName, sName, sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
    RegReadStringValue = (RetVal = 0)
End Function 'RegReadStringValue
'=======================================================================================================

Function RegReadExpStringValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal

    RetVal = oReg.GetExpandedStringValue(hDefKey, sSubKeyName, sName, sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetExpandedStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
    RegReadExpStringValue = (RetVal = 0)
End Function 'RegReadExpStringValue
'=======================================================================================================

Function RegReadMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
    Dim RetVal

    RetVal = oReg.GetMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetMultiStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, arrValues)
    RegReadMultiStringValue = (RetVal = 0 AND IsArray(arrValues))
End Function 'RegReadMultiStringValue
'=======================================================================================================

Function RegReadDWordValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal

    RetVal = oReg.GetDWORDValue(hDefKey, sSubKeyName, sName, sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetDWORDValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
    RegReadDWordValue = (RetVal = 0)
End Function 'RegReadDWordValue
'=======================================================================================================

Function RegReadBinaryValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal

    RetVal = oReg.GetBinaryValue(hDefKey, sSubKeyName, sName, sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetBinaryValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
    RegReadBinaryValue = (RetVal = 0)
End Function 'RegReadBinaryValue
'=======================================================================================================

Function RegReadQWordValue(hDefKey, sSubKeyName, sName, sValue)
    Dim RetVal

    RetVal = oReg.GetQWORDValue(hDefKey, sSubKeyName, sName, sValue)
    If Not RetVal = 0 AND f64 Then RetVal = oReg.GetQWORDValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, sValue)
    RegReadQWordValue = (RetVal = 0)
End Function 'RegReadQWordValue
'=======================================================================================================

'Enumerate a registry key to return all values
Function RegEnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    If f64 Then
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames32, arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrNames64, arrTypes64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32, "\") & "\" & Join(arrNames64, "\")), "\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32, "\") & "\" & Join(arrTypes64, "\")), "\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    End If 'f64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'=======================================================================================================

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey, sSubKeyName, arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    If f64 Then
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrKeys64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32, "\") & "\" & Join(arrKeys64, "\")), "\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys)
    End If 'f64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'=======================================================================================================

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos
    Dim sKey, sVer
    Dim fReplaced

    fReplaced = False
    For Each sVer in dicActiveC2RVersions.Keys
        sKey = REG_OFFICE & sVer & REG_C2RVIRT_HKLM
        If InStr(sSubKeyName, sKey) > 0 Then
            sSubKeyName = Replace(sSubKeyName, sKey, "")
            fReplaced = True
            Exit For
        End If
    Next
    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName, 17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - 17)
            Else
                iPos = InStr(sSubKeyName, "\")
                Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) -iPos)
            End If
        Case HKLM
            If Left(sSubKeyName, 17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) - 17)
            Else
                iPos = InStr(sSubKeyName, "\")
                Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName) -iPos)
            End If
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
    End Select 'hDefKey
    If fReplaced Then
        sSubKeyName = sKey & sSubKeyName
        Wow64Key = sKey & Wow64Key
    End If
End Function 'Wow64Key
'=======================================================================================================

'64 bit aware wrapper to return the requested folder 
Function GetFolderPath(sPath)
    GetFolderPath = True
    If oFso.FolderExists(sPath) Then Exit Function
    If f64 AND oFso.FolderExists(Wow64Folder(sPath)) Then
        sPath = Wow64Folder(sPath)
        Exit Function
    End If
    GetFolderPath = False
End Function 'GetFolderPath
'=======================================================================================================

'Enumerates subfolder names of a folder and returns True if subfolders exist
Function EnumFolderNames (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Name & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders) > 0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders, Len(sSubFolders) - 1), ","))
    EnumFolderNames = Len(sSubFolders) > 0
End Function 'EnumFolderNames
'=======================================================================================================

'Enumerates subfolders of a folder and returns True if subfolders exist
Function EnumFolders (sFolder, arrSubFolders)
    Dim Folder, Subfolder
    Dim sSubFolders
    
    If oFso.FolderExists(sFolder) Then
        Set Folder = oFso.GetFolder(sFolder)
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If f64 AND oFso.FolderExists(Wow64Folder(sFolder)) Then
        Set Folder = oFso.GetFolder(Wow64Folder(sFolder))
        For Each Subfolder in Folder.Subfolders
            sSubFolders = sSubFolders & Subfolder.Path & ","
        Next 'Subfolder
    End If
    If Len(sSubFolders) > 0 Then arrSubFolders = RemoveDuplicates(Split(Left(sSubFolders, Len(sSubFolders) - 1), ","))
    EnumFolders = Len(sSubFolders) > 0
End Function 'EnumFolders
'=======================================================================================================

Sub GetFolderStructure (Folder)
    Dim SubFolder
    
    For Each SubFolder in Folder.SubFolders
        ReDim Preserve arrMseFolders(UBound(arrMseFolders) + 1)
        arrMseFolders(UBound(arrMseFolders)) = SubFolder.Path
        GetFolderStructure SubFolder
    Next 'SubFolder
End Sub 'GetFolderStructure
'=======================================================================================================

'Remove duplicate entries from a one dimensional array
Function RemoveDuplicates(Array)
    Dim Item
    Dim oDic
    
    Set oDic = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not oDic.Exists(Item) Then oDic.Add Item, Item
    Next 'Item
    RemoveDuplicates = oDic.Keys
End Function 'RemoveDuplicates
'=======================================================================================================

'Identify user SID's on the system
Function GetUserSids(sContext)
    Dim oAccount
    Dim sValue
    Dim i, n, iRetVal
    Dim arrKeys
    Dim key
    On Error Resume Next

    sUserName = oShell.ExpandEnvironmentStrings ("%USERNAME%") 
    sDomain = oShell.ExpandEnvironmentStrings ("%USERDOMAIN%") 

    ReDim arrKeys(-1)
    Select Case sContext
    Case "Current"
        ' Need to suppress errors for this
        sCurUserSid = ""
        If RegEnumKey(HKCU, "Software\Microsoft\Protected Storage System Provider\", arrKeys) Then
            sCurUserSid = arrKeys(0)
        ElseIf RegReadValue (HKCU, "Software\Microsoft\Windows\CurrentVersion\FileAssociations\","UserSid", sValue, "REG_SZ") Then
            sCurUserSid = sValue
        ElseIf RegReadValue (HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\","LastLoggedOnUserSID", sValue, "REG_SZ") Then
            sCurUserSid = sValue
        Else
            If RegEnumKey(HKEY_USERS, "", arrKeys) Then
                For Each key in arrKeys
                    If LCase(sUserName) = LCase(oWmiLocal.Get("Win32_SID.SID='" & key & "'").AccountName) Then sCurUserSid = key
                Next 'key
            End If
            If sCurUserSid = "" Then
                Set oAccount = oWMILocal.Get ("Win32_UserAccount.Name='" & sUserName & "', Domain='" & sDomain & "'") 
                sCurUserSid = oAccount.SID
            End If
        End If 'RegEnumKey
    
    Case "UserUnmanaged"
        If RegEnumKey(HKLM, REG_GLOBALCONFIG, arrKeys) Then
            n = 0
            For i = 0 To UBound(arrKeys)
                If Len(arrKeys(i)) > 8 Then
                    Redim Preserve arrUUSids(n)
                    arrUUSids(n) = arrKeys(i)
                    n = n + 1
                End If 'Len(arrKeys)
            Next 'i
        End If 'RegEnumKey
    Case "UserManaged"
        If RegEnumKey(HKLM, REG_CONTEXTUSERMANAGED, arrKeys) Then
            n = 0
            For i = 0 To UBound(arrKeys)
                If Len(arrKeys(i)) > 8 Then
                    Redim Preserve arrUMSids(n)
                    arrUMSids(n) = arrKeys(i)
                    n = n + 1
                End If 'Len(arrKeys)
            Next 'i
        End If 'RegEnumKey
    Case Else
    End Select
End Function
'=======================================================================================================

Function GetArrayPosition (Arr, sProductCode)
    Dim iPos
    On Error Resume Next

    GetArrayPosition = -1
    If CheckArray(Arr) And (IsValidGuid(sProductCode, GUID_UNCOMPRESSED) OR fGuidCaseWarningOnly) Then
        For iPos = 0 To UBound(Arr)
            If Arr(iPos, COL_PRODUCTCODE) = sProductCode Then 
                GetArrayPosition = iPos
                Exit For
            End If
        Next 'iPos
    End If 'CheckArray
    If iPos = -1 Then WriteDebug sActiveSub, "Warning: Invalid ArrayPosition for " & sProductCode & " - Stack: " & sStack
End Function
'=======================================================================================================

Function GetArrayPositionFromPattern (Arr, sProductCodePattern)
    Dim iPos
    On Error Resume Next

    GetArrayPositionFromPattern = -1
    If CheckArray(Arr) Then
        For iPos = 0 To UBound(Arr)
            If InStr(Arr(iPos, COL_PRODUCTCODE), sProductCodePattern) > 0 Then 
                GetArrayPositionFromPattern = iPos
                Exit For
            End If
        Next 'iPos
    End If 'CheckArray
End Function
'=======================================================================================================

Sub InitMasterArray
    On Error Resume Next
    
    If CheckArray(arrAllProducts) Then iPCount = iPCount + (UBound(arrAllProducts) + 1)
    If CheckArray(arrMProducts) Then iPCount = iPCount + (UBound(arrMProducts) + 1)
    If CheckArray(arrUUProducts) Then iPCount = iPCount + (UBound(arrUUProducts) + 1)
    If CheckArray(arrUMProducts) Then iPCount = iPCount + (UBound(arrUMProducts) + 1)
    If CheckArray(arrMVProducts) Then iPCount = iPCount + (UBound(arrMVProducts) + 1 )
    
    ReDim arrMaster(iPCount- 1, UBOUND_MASTER)
    
    If CheckArray(arrAllProducts) Then CopyToMaster arrAllProducts
    If CheckArray(arrMProducts) Then CopyToMaster arrMProducts
    If CheckArray(arrUUProducts) Then CopyToMaster arrUUProducts
    If CheckArray(arrUMProducts) Then CopyToMaster arrUMProducts
    If CheckArray(arrMVProducts) Then CopyToMaster arrMVProducts
End Sub
'=======================================================================================================

Function CheckArray(Arr)
    Dim sTmp
    Dim iRetVal
    On Error Resume Next
    
    CheckArray = True
    If Not IsArray(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If IsNull(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If IsEmpty(Arr) Then 
        CheckArray = False
        Exit Function
    End If
    If UBound(Arr) = -1 Then 
        If Not Err = 0 Then 
            Err.Clear
            Redim Arr(-1)
        End If 'Err
        CheckArray = False
        Exit Function
    End If
End Function
 
'=======================================================================================================

Function CheckObject(Obj)
    On Error Resume Next
    
    CheckObject = True
    If Not IsObject(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If IsNull(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If IsEmpty(Obj) Then 
        CheckObject = False
        Exit Function
    End If
    If Not Obj.Count > 0 Then 
        If Not Err = 0 Then
            Err.Clear
            Set Obj = Nothing
        End If 'Err
        CheckObject = False
        Exit Function
    End If
End Function
'=======================================================================================================

Sub CheckError(sModule, sErrorHandler)
    Dim sErr
    If Not Err = 0 Then 
        sErr = GetErrorDescription(Err)
        If Not sErrorHandler = "" Then 
            sErrorHandler = sModule & sErrorHandler
            ErrorRelay(sErrorHandler)
        End If
    End If 'Err = 0
    Err.Clear
End Sub
'=======================================================================================================

Sub ErrorRelay(sErrorHandler)
    Select Case (sErrorHandler)
        Case "CheckPreReq_ErrorHandler" : CheckPreReq_ErrorHandler
        Case "FindAllProducts_ErrorHandler3x" : FindAllProducts_ErrorHandler3x
        Case "FindProducts1","FindProducts2","FindProducts4" : FindProducts_ErrorHandler Int(Right(sErrorHandler, 1))
        Case Else
    End Select
End Sub
'=======================================================================================================

Function GetErrorDescription (Err)
    If Not Err = 0 Then 
        GetErrorDescription = "Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                              "; Err# (Dec): " & Err & "; Description : " & Err.Description & _
                              "; Stack: " & sStack
    End If 'Err = 0
End Function
'=======================================================================================================

Sub ParseCmdLine
    Dim iCnt, iArgCnt
    
    If fLogFull Then
        fLogChainedDetails = True
        fFileInventory = True
        fFeatureTree = True
    End If 'fLogFull
    
    If fLogVerbose Then
        fLogChainedDetails = True
        fFeatureTree = True
    End If 'fLogVerbose
    
    iArgCnt = Wscript.Arguments.Count
    If iArgCnt>0 Then
        For iCnt = 0 To iArgCnt- 1
            Select Case UCase(Wscript.Arguments(iCnt))
            
            Case "/A","-A","/ALL","-ALL","ALL"
                fListNonOfficeProducts = True
            Case "/BASIC","-BASIC","BASIC"
                fBasicMode = True
            Case "/DCS","-DCS","/DISALLOWCSCRIPT","-DISALLOWCSCRIPT","DISALLOWCSCRIPT"
                fDisallowCScript = True
            Case "/F","-F","/FULL","-FULL","FULL"
                fLogFull = True
                fLogChainedDetails = True
                fFileInventory = True
                fFeatureTree = True
            Case "/FI","-FI","/FILEINVENTORY","-FILEINVENTORY","FILEINVENTORY"
                fFileInventory = True
            Case "/FT","-FT","/FEATURETREE","-FEATURETREE","FEATURETREE"
                fFeatureTree = True
            Case "/L","-L","/LOGFOLDER","-LOGFOLDER","LOGFOLDER"
                If iArgCnt > iCnt Then sPathOutputFolder = Wscript.Arguments(iCnt + 1)
            Case "/LD","-LD","/LOGCHAINEDDETAILS","-LOGCHAINEDDETAILS"
                fLogChainedDetails = True
            Case "/LV","-LV","/L*V","-L*V","/LOGVERBOSE","-LOGVERBOSE","/VERBOSE","-VERBOSE"
                fLogVerbose = True
                fLogChainedDetails = True
                fFeatureTree = True
            Case "/Q","-Q","/QUIET","-QUIET","QUIET"
                fQuiet = True
            Case "/?","-?","?"
                Wscript.Echo vbCrLf & _
                 "ROIScan (Robust Office Inventory Scan) - Version " & SCRIPTBUILD & vbCrLf & _
                 "Copyright (c) 2008, 2009, 2010 Microsoft Corporation. All Rights Reserved." & vbCrLf & vbCrLf & _
                 "Inventory tool for to create a log of " & vbCrLf & "installed Microsoft Office applications." & vbCrLf & _
                 "Supports Office 2000, XP, 2003, 2007, 2010, 2013, 2016, O365 " & vbCrLf & vbCrLf & _
                 "Usage:" & vbTab & "ROIScan.vbs [Options]" & vbCrLf & vbCrLf & _
                 " /?" & vbTab & vbTab & vbTab & "Display this help"& vbCrLf &_
                 " /All" & vbTab & vbTab & vbTab & "Include non Office products" & vbCrLf &_
                 " /Basic" & vbTab & vbTab & vbTab & "Log less product details" & vbCrLf & _
                 " /Full" & vbTab & vbTab & vbTab & "Log all additional product details" & vbCrLf & _
                 " /LogVerbose" & vbTab & vbTab & "Log additional product details" & vbCrLf & _
                 " /LogChainedDetails" & vbTab & "Log details for chained .msi packages" & vbCrLf & _
                 " /FeatureTree" & vbTab & vbTab & "Include treeview like feature listing" & vbCrLf & _
                 " /FileInventory" & vbTab & vbTab & "Scan all file version" & vbCrLf & _
                 " /Logfolder <Path>" & vbTab & "Custom directory for log file" & vbCrLf &_
                 " /Quiet" & vbTab & vbTab & vbTab & "Don't open log when done"
                Wscript.Quit
            Case Else
            End Select
        Next 'iCnt
    End If 'iArgCnt>0
End Sub
'=======================================================================================================

Sub CleanUp
    On Error Resume Next
    Set oReg = Nothing
    TextStream.Close
    Set TextStream = Nothing
    Set oFso = Nothing
    Set oMsi = Nothing
    Set oShell = Nothing
End Sub
'=======================================================================================================
