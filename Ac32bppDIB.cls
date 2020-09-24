VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ac32bppDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Credits/Acknowledgements - Thanx goes to:
'   Paul Caton for his class on calling non VB-Friendly DLLs that use _cdecl calling convention
'       Used when calling non VB-friendly zLIB dll versions
'   Alfred Koppold for his PNG, VB-only, decompression routines.
'       Used when zLib & GDI+ not available
'   Carles P.V for his pvResize logic
'       Used when manually scaling images with NearestNeighbor or BiLinear interpolation
'   www.zlib.net for their free zLIB.dll, the standard DLL for compressing/decompressing PNGs
'       Without it, we'd be limited to GDI+ for creating PNGs
'   coders like you that provide constructive criticism to make this class better & more all-inclusive
'       Without your comments, this project probably would have died several versions/updates ago
' For most current updates/enhancements visit the following:
'   Visit http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                    O V E R V I E W
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' About 32bpp pre-multiplied RGB (pARGB) bitmaps, if you are not aware.
'   - These are used specifically for the AlphaBlend API & are GDI+ compatible
'   Advantages:
'       - Images can be per-pixel alpha blended
'       - Opacity can be simultaneously adjusted during rendering
'       - AlphaBlend does both BitBlt & StretchBlt for pARGB images.
'       - Speed: AlphaBlend & GDI+ are pretty quick APIs vs manual blending
'   Disadvantages:
'       - The original RGB values are permanently destroyed during pre-multiplying
'           -- Premultiplied formula: preMultipliedRed=(OriginalRed * Alpha) \ 255
'           -- There is no way to convert pARGB back to non-premultiplied RGB values
'              The formula would be: reconstructedRed=(preMultipliedRed * 255) \ Alpha.
'               but because of integer division when pre-multiplying, the result is only
'               close and if this should be premultiplied again & converted again, the
'               alphas can get more transparent with every iteration.
'               Fully opaque pixels & fully transparent pixels are not affected.
'           ** Note: When images are converted to PNG formats, removal of
'              premultiplication is performed to meet PNG specs.
'       - Displaying a pre-multiplied bitmap without AlphaBlend will not result in
'           the image being displayed as expected.
'       - Not ideal for saving due to its size: SizeOf= W x H x 4
'           -- better to save source image instead or compress the DIB bytes using favorite compression utility
'           -- with GDI+ or zLib, image can be converted to PNG for storage
'       - AlphaBlend (msimg32.dll) is not included/compatible with Win95, NT4 and lower
'       - AlphaBlend on Win9x systems can be buggy, especially when rendering to DIBs vs DDBs
' Note that GDI+ is standard on WinXP+, and can be used on Win98,ME,2K, & on NT4 if SP6 is installed
' Download GDI+ from:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdicpp/GDIPlus/GDIPlus.asp

' ----------------------------------------------
' About Win95, Win98, NT3.5, NT4 & WinME support
' ----------------------------------------------
' The routines will not honor AlphaBlend if it exists on those systems. Win98's version,
' for example, has several bugs that can crash the application when AlphaBlending to DIBs.
' NT4, NT3.5 & Win95 do not come with AlphaBlend and I do not have WinME to test with.
' Therefore, to support these systems, the Render routine will alphablend manually
' regardless if the AlhpaBlend API (msimg32.dll) exists on the system or not.
' However, this can be overridden by you. See isAlphaBlendFriendly routine


' Class Purpose:
' ----------------------------------------------
' This class holds the 32bpp image. It also marshals any new image thru
' the battery of parsers to determine best method for converting the image
' to a 32bpp alpha-compatible image. It handles rendering, rotating, scaling,
' mirroring of DIBs using manual processes, AlphaBlend, and/or GDI+.

' The parser order is very important for fastest/best results...
' AcPNGparser :: will convert PNG, all bit depths; aborts quickly if not PNG
' AcGIFparser :: will convert non-transparent/transparent GIFs; aborts quickly
' cICOpraser :: will convert XP-Alpha, paletted, true color, & Vista PNG icons
'               -- can also convert most non-animated cursors
' AcBMPparser :: will convert bitmaps, wmf/emf & jpgs

' The parsers are efficient. Most image formats have a magic number that give
'   a hint to what type of image the file/stream is. However, checks need to
'   be employed because non-image files could feasibly have those same magic
'   numbers. If the image is determined not to be one the parser is designed
'   to handle, the parser rejects it and the next parser takes over.  The
'   icon parser is slightly different because PNG files can be included into
'   a Vista ico file. When this occurs, the icon parser will pass off the
'   PNG format to the PNG parser automatically.
' And last but not least, the parsers have no advanced knowledge of the image
' format; as far as they are concerned, anything passed is just a byte array

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                       CHANGE HISTORY
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Accompanying FAQ.rtf is updated with every change
' Last changed: 11 Apr 07. See change history within the FAQ file
' 26 Dec 06: First version
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflciting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take liberties, changing parameter types, in several APIs throughout

' Used to determine operating system
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Type OSVERSIONINFOEX
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128 ' up to here is OSVERSIONINFO vs EX
   wServicePackMajor As Integer ' 8 bytes larger than OSVERSIONINFO
   wServicePackMinor As Integer
   wSuiteMask As Integer
   wProductType As Byte
   wReserved As Byte
End Type

' APIs used to manage the 32bpp DIB
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef Pointer As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const STRETCH_HALFTONE As Long = &H4&
Private Const OBJ_BITMAP As Long = &H7&
Private Const OBJ_METAFILE As Long = &H9&
Private Const OBJ_ENHMETAFILE As Long = &HD&

' APIs used to create files
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1&

' ////////////////////////////////////////////////////////////////
' Unicode-capable Drag and Drop of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, _
    ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, _
    ByVal paCNT As Long, ByRef paTypes As Integer, _
    ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (lpString As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

' ////////////////////////////////////////////////////////////////
' Unicode-capable Pasting of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
' ////////////////////////////////////////////////////////////////
Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type
Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type
Private Type STGMEDIUM
    TYMED As Long
    Data As Long
    pUnkForRelease As IUnknown
End Type


' used to create the checkerboard pattern on demand
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' used when saving an image or part of the image
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' reusable UDT for 1 & 2 dim arrays
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette As Long
End Type

Private Const AC_SRC_OVER = &H0&
Private Const AC_SRC_ALPHA = &H1&

Public Enum AeImageFormat    ' source image format
    AimgError = -1  ' no DIB has been initialized
    AimgNone = 0    ' no image loaded
    AimgBitmap = 1  ' standard bitmap or jpg
    AimgIcon = 3    ' standard icon
    AimgWMF = 2     ' windows meta file
    AimgEMF = 4     ' enhanced WMF
    AimgCursor = 5  ' standard cursor
    AimgBmpARGB = 6  ' 32bpp bitmap where RGB is not pre-multiplied
    AimgBmpPARGB = 7 ' 32bpp bitmap where RGB is pre-multiplied
    AimgIconARGB = 8 ' XP-type icon; 32bpp ARGB
    AimgGIF = 9      ' gif; if class.Alpha=True, then transparent GIF
    AimgPNG = 10     ' PNG image
    AimgPNGicon = 11 ' PNG in icon file (Vista)
    AimgCursorARGB = 12 ' alpha blended cursors? do they exist yet?
    AimgCheckerBoard = 64 ' image is displaying own checkerboard pattern; no true image
End Enum

Public Enum AePngProperties ' following are recognized "Captions" within a PNG file
    AtxtTitle = 1           ' See AcPNGwriter.SetPngProperty for more information
    AtxtAuthor = 2
    AtxtDescription = 4
    AtxtCopyright = 8
    AtxtCreationTime = 16
    AtxtSoftware = 32
    AtxtDisclaimer = 64
    AtxtWarning = 128
    AtxtSource = 256
    AtxtComment = 512
    ' special properties
    AtxtLargeBlockText = 1024 ' this is free-form text can be of any length & contain most any characters
    AdateTimeModified = 2048  ' date/time of the last image modification (not the time of initial image creation)
    AcolorDefaultBkg = 4096   ' default background color to use if PNG viewer does not do transparency
    AfilterType = 8192        ' one of the AeFilterMethods values
    ClearAllProperties = -1  ' resets all PNG properties
End Enum

Public Enum AeTrimOptions    ' see TrimImage method
    trimAll = 0             ' can be combined using OR
    trimLeft = 1
    trimTop = 2
    trimRight = 4
    trimBottom = 8
End Enum

Public Enum eScaleOptions   ' See ScaleImage method
    scaleToSize = 0         ' [Default] will always scale
    scaleDownAsNeeded = 1   ' will only scale down if image won't fit
    ScaleStretch = 2        ' wll always stretch/distort
End Enum

Public Enum AeGrayScaleFormulas
    AgsclNTSCPAL = 0     ' R=R*.299, G=G*.587, B=B*.114 - Default
    AgsclCCIR709 = 1     ' R=R*.213, G=G*.715, B=B*.072
    AgsclSimpleAvg = 2   ' R,G, and B = (R+G+B)/3
    AgsclRedMask = 3     ' uses only the Red sample value: RGB = Red / 3
    AgsclGreenMask = 4   ' uses only the Green sample value: RGB = Green / 3
    AgsclBlueMask = 5    ' uses only the Blue sample value: RGB = Blue / 3
    AgsclRedGreenMask = 6 ' uses Red & Green sample value: RGB = (Red+Green) / 2
    AgsclBlueGreenMask = 7 ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    AgsclNone = -1
End Enum

Public Enum AeFilterMethods
    AfilterDefault = 0     ' paletted PNGs will use AfilterNone while others will use AfilterPaeth
    AfilterNone = 1        ' no byte preparation used; else preps bytes using one of the following
    AfilterAdjLeft = 2     ' see AcPNGwriter.EncodeFilter_Sub
    AfilterAdjTop = 3      ' see AcPNGwriter.EncodeFilter_Up
    AfilterAdjAvg = 4      ' see AcPNGwriter.EncodeFilter_Avg
    AfilterPaeth = 5       ' see AcPNGwriter.EncodeFilter_Paeth
    AfilterAdaptive = 6    ' this is a best guess of the above 4 (can be different for each DIB scanline)
End Enum

Public Enum AeRegionStyles     ' See CreateRegion
    AregionBounds = 0
    AregionEnclosed = 1
    AregionShaped = 2
End Enum

Public Enum AeConstants      ' See SourceIconSizes
    ATRUE_COLOR = &HFF000000
    AHIGH_COLOR = &HFFFF00
    ATRUE_COLOR_ALPHA = &HFFFFFFFF
End Enum

'Private m_PNGprops As AcPNGwriter    ' used for more advanced PNG creation options
Private m_StretchQuality As Boolean ' if true will use BiLinear or better interpolation
Private m_Handle As Long        ' handle to 32bpp DIB
Private m_Pointer As Long       ' pointer to DIB bits
Private m_Height As Long        ' height of DIB
Private m_Width As Long         ' width of DIB
Private m_hDC As Long           ' DC if self-managing one
Private m_prevObj As Long       ' object deselected from DC when needed
Private m_osCAP As Long         ' See Class_Initialize
Private m_Format As AeImageFormat ' type of source image
Private m_ManageDC As Boolean   ' does class manage its own DC
Private m_AlphaImage As Boolean ' does the DIB contain alpha/transparency
Private m_GDItoken As Long
Private m_ImageByteCache() As Byte  ' should you want the DIB class to cache original bytes
' ^^ N/A if image is loaded by handle, stdPicture, or resource

Public Function LoadPicture_File(ByVal FileName As String, Optional ByVal iconCx As Long, Optional ByVal iconCy As Long, Optional ByVal SaveFormat As Boolean, Optional ByVal iconBitDepth As Long = 32) As Boolean

    ' PURPOSE: Convert passed image file into a 32bpp image
    
    ' Parameters.
    ' FileName :: full path of file. Validation occurs before we continue
    ' iconCx :: desired width of icon if file is an icon file. Default is 32x32
    ' iconCy :: desired height of icon if file is an icon file. Default is 32x32
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    ' iconBitDepth :: the desired bit depth of an icon if the resource is an icon file
    
    ' Why would you want to save the bytes? If this is being used in a usercontrol,
    ' saving the bytes will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.

    On Error Resume Next
    Dim hFile As Long
    
    hFile = AiparseGetFileHandle(FileName, True, ((m_osCAP And 24) = 8))
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    
    If GetFileSize(hFile, 0&) > 56 Then
        
        ' no image file/stream can be less than 57 bytes and still be an image
        Dim aDIB() As Byte  ' dummy array
        LoadPicture_File = LoadPictureEx(hFile, FileName, aDIB(), iconCx, iconCy, 0&, 0&, SaveFormat, iconBitDepth)
    
    End If
    CloseHandle hFile
    
End Function

Friend Property Let Alpha(isAlpha As Boolean)
    m_AlphaImage = isAlpha  ' determines the flags used for AlphaBlend API
    ' this flag is set by the various image parsers; setting it yourself
    ' can produce less than desirable effects.
    ' Used in Me.Render & Me.TrimImage, AcPNGwriter.OptimizeTrueColor & AcPNGwriter.PalettizeImage
End Property
Public Property Get Alpha() As Boolean
    Alpha = m_AlphaImage
End Property

Public Property Let HighQualityInterpolation(Value As Boolean)
    ' When possible GDI+ will be used for stretching & rotation.
    ' If GDI+ is used,then high quality equates to BiCubic interpolation
    ' If not used, then BiLinear (manual processing) will be used.
    ' If High Quality is false, then Nearest Neighbor (very fast) interpolation used
    m_StretchQuality = Value
End Property
Public Property Get HighQualityInterpolation() As Boolean
    HighQualityInterpolation = m_StretchQuality
End Property

Public Property Get ImageType() As AeImageFormat
    ImageType = m_Format    ' returns image format of the source image
End Property
Friend Property Let ImageType(iType As AeImageFormat)
    m_Format = iType    ' set by the various image parsers. This is not used
    ' anywhere in these classes, you can do with it what you want -- for now.
End Property

Public Property Get Width() As Long
    Width = m_Width     ' width of image in pixels
End Property
Public Property Get Height() As Long
    Height = m_Height   ' height of image in pixels
End Property
Public Property Get BitsPointer() As Long
    BitsPointer = m_Pointer ' pointer to the bits of the image
End Property
Public Property Get scanWidth() As Long
    scanWidth = m_Width * 4&    ' number of bytes per scan line
End Property
Public Property Get Handle() As Long
    Handle = m_Handle   ' the picture handle of the image
End Property

Public Function LoadDIBinDC(ByVal bLoad As Boolean) As Long

    ' Purpose: Select/Unselect the DIB into a DC.
    ' Returns the DC handle when image is loaded
    ' Called by image parser if it needs to paint the image into the DIB
       
    If bLoad = True Then
        Dim tDC As Long
        If Not m_Handle = 0& Then    ' do we have an image?
            If m_hDC = 0& Then        ' do we have a DC?
                tDC = GetDC(0&)     ' if not create one
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
            If m_prevObj = 0& Then
                m_prevObj = SelectObject(m_hDC, m_Handle)
            End If
            LoadDIBinDC = m_hDC
        End If
    Else
        If Not m_prevObj = 0& Then
            SelectObject m_hDC, m_prevObj
            If m_ManageDC = False Then
                DeleteObject m_hDC
                m_hDC = 0&
            End If
            m_prevObj = 0&
        End If
    End If
End Function

Public Property Let ManageOwnDC(bManage As Boolean)
    ' Determines whether or not this class will manage its own DC
    ' If false, then a DC is created each time the image needs to be Rendered
    Dim tDC As Long
    If bManage = False Then     ' removing management of DC
        If Not m_hDC = 0& Then   ' DC does exist, destroy it
            ' first remove the dib, if one exists
            If Not m_Handle = 0& Then SelectObject m_hDC, m_prevObj
            m_prevObj = 0&
        End If
        DeleteDC m_hDC
        m_hDC = 0&
    Else                        ' allowing creation of dc
        If m_hDC = 0& Then        ' create DC only if we have a dib to put in it
            If Not m_Handle = 0& Then
                tDC = GetDC(0&)
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
        End If
    End If
    m_ManageDC = bManage
End Property
Public Property Get ManageOwnDC() As Boolean
    ManageOwnDC = m_ManageDC
End Property

Public Property Get isAlphaBlendFriendly() As Boolean
    isAlphaBlendFriendly = ((m_osCAP And 1) = 1)
    ' WinNT4 & below and Win95 are not shipped with msimg32.dll (AlphaBlend API)
    ' Win98 has bugs & would believe that WinME is buggy too but don't know for sure
    ' Therefore, the Rendering in this class will not use AlphaBlend on these
    ' operating systems even if the DLL exists, but will use GDI+ if available
    ' Can be overridden by setting this property to True
End Property
Public Property Let isAlphaBlendFriendly(Enabled As Boolean)
    ' This has been provided to override safety of using AlphaBlend on Win9x systems.
    ' Caution. Only set this when rendering to a known device dependent bitmap (DDB)
    ' Alphablend can crash when rendering DIB to DIB vs DIB to DDB. Be warned.
    If Enabled = True Then
        ' Overriding in play: allow AlphaBlend if system is Win98 or better
        ' By default this is already set for Win2K or better
        If ((m_osCAP And 8) = 8) Then m_osCAP = m_osCAP Or 1
    Else
        m_osCAP = m_osCAP And Not 1 ' disallow AlphaBlend
    End If
End Property
Public Property Get isGDIplusEnabled() As Boolean
    ' identifies if GDI+ is usable on the system.
    ' Before this property is set, GDI+ is tested to ensure it is usable
    isGDIplusEnabled = ((m_osCAP And 2) = 2)
End Property
Public Property Let isGDIplusEnabled(Enabled As Boolean)
    ' Sets the property. If set to False by you, GDI+ will not be used
    ' for any rendering, but still may be used to create PNG files if needed
    
    If Not Enabled = Me.isGDIplusEnabled Then
        m_osCAP = (m_osCAP And Not 2)
        If Enabled Then
            If (m_osCAP And 32) = 0 Then ' else Win95, NT4 SP5 or lower
                Dim cGDIp As New AcGDIPlus
                If cGDIp.isGDIplusOk() = True Then m_osCAP = m_osCAP Or 2
            End If
        End If
    End If
End Property



Public Function InitializeDIB(ByVal Width As Long, ByVal Height As Long) As Boolean

    ' Creates a blank (all black, all transparent) DIB of requested height & width
    
    Dim tBMPI As BITMAPINFO, tDC As Long
    
    DestroyDIB ' clear any pre-existing dib
    
    If Width < 0& Then Exit Function
    If Height = 0& Then
        Exit Function
    ElseIf Height < 0& Then
        Height = Abs(Height) ' no top-down dibs
    End If
    
    On Error Resume Next
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40&
        .biSizeImage = .biHeight * .biWidth * 4&
    End With
    If Err Then
        Err.Clear
        ' only possible error would be that Width*Height*4& is absolutely huge
        Exit Function
    End If
    
    tDC = GetDC(0&) ' get screen DC
    m_Handle = CreateDIBSection(tDC, tBMPI, 0&, m_Pointer, 0&, 0&)
    If m_ManageDC = True Then
        ' create a DC if class is managing its own & one isn't created yet
        If m_hDC = 0& Then m_hDC = CreateCompatibleDC(tDC)
    End If
    ' release the screen DC if we captured it
    ReleaseDC 0&, tDC
    
    If Not m_Handle = 0& Then    ' let's hope system resources allowed DIB creation
        m_Width = Width
        m_Height = Height
        m_AlphaImage = True
        m_Format = AimgNone
        InitializeDIB = True
    End If

End Function

Public Sub DestroyDIB()
    
    ' PURPOSE: Destroy any existing image
    If Not m_hDC = 0& Then   ' do we have a DC?
        ' do we have an image; if so get it out of the DC
        If Not m_prevObj = 0& Then SelectObject m_hDC, m_prevObj
        ' destroy our DC, no point in keeping it w/o image
        DeleteObject m_hDC
        m_hDC = 0&
    End If
    ' if we do have an image, destroy it now
    If Not m_Handle = 0& Then
        DeleteObject m_Handle
        Erase m_ImageByteCache
    End If
    ' reset other image attributes
    m_Width = 0&
    m_Height = 0&
    m_Handle = 0&
    m_Pointer = 0&
    m_prevObj = 0&
    m_AlphaImage = False
    m_Format = AimgError
End Sub

Public Sub EraseDIB()
    ' Function clears out an existing DIB, making it 100% transparent/black
    If Not m_Handle = 0& Then
        FillMemory ByVal m_Pointer, m_Width * m_Height * 4&, 0
        m_Format = AimgNone
        m_AlphaImage = True
    End If
    
End Sub

Public Function Render(ByVal destinationDC As Long, Optional ByVal destX As Long, Optional ByVal destY As Long, Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, Optional ByVal Opacity As Long = 100&, Optional ByVal Blend As Boolean = True, Optional ByVal SetHalfTone As Boolean = True, Optional ByRef destHostDIB As Ac32bppDIB, Optional ByVal grayScale As AeGrayScaleFormulas = AgsclNone) As Boolean

    ' PURPOSE: Render an existing 32bpp DIB to a target DC
    
    ' Parameters.
    ' destinationDC :: target DC to draw to
    ' destX, destY :: the top/left coordinates to draw to, default is 0,0
    ' destWidth, destHeight :: the width and height to draw to, default is the image's width & height
    ' srcX, srcY :: the left & top offset within the DIB
    ' srcWidth, srcHeight :: the amount of DIB to be rendered
    ' Opacity :: how opaque to draw the image, default is 100% opaque
    ' Blend :: no longer used, reserved
    ' SetHalfTone :: if True, then the destination DC's stretch mode will be modified to
    '       produce better quality results. This option is not available on Win9x systems.
    '       Tip: When AlphaBlending to another DIB set to False
    '            When AlphaBlending to CompatibleBitmap (DDB) or visible DC set to True
    ' destHostDIB :: When rendering from DIB class to DIB class, pass the destination
    '       DIB class to ensure alpha blending occurs correctly on systems that do not
    '       support GDI+ or AlphaBlend APIs
    ' grayscale :: one of several formulas to grayscale while rendering (optional)

    Dim lBlendFunc As Long, tDC As Long, hOldImage As Long
    Dim lStretchMode As Long
    Dim aResizedBytes() As Byte, aMirrorBytes() As Byte
    Dim bStretching As Boolean
    Dim bMirroring As Boolean
    Dim bCanUseAlphaBlend As Boolean
    
    ' validate a few things
    If m_Handle = 0& Then
        Exit Function
    ElseIf destinationDC = 0& Then
        Exit Function
    ElseIf srcWidth < 0& Then   ' AlphaBlend is not compatible with negative width/height
        Exit Function          ' negative values used in APIs like StretchBlt for mirroring
    ElseIf srcHeight < 0& Then
        Exit Function
    End If
    
    If Opacity = 0& Then
        Render = True
        Exit Function   ' pointless if image is 100% transparent
    Else
        Opacity = Abs(Opacity) Mod 100
        If Opacity = 0& Then Opacity = 100&
    End If
    
    ' validate optional parameters for source image
        If srcWidth = 0& Then srcWidth = m_Width
        If srcHeight = 0& Then srcHeight = m_Height
        If SrcX < 0& Then SrcX = 0&  ' source X,Y cannot be negative
        If SrcY < 0& Then SrcY = 0&  ' but the dest X,Y can be
    ' AlphaBlend requires that the source rectangle fit within the image
        If SrcX + srcWidth > m_Width Then srcWidth = m_Width - SrcX
        If SrcY + srcHeight > m_Height Then srcHeight = m_Height - SrcY
    ' validate optional parameters for destination image
        If destWidth = 0& Then
            destWidth = m_Width
        Else
            If destWidth < 0& Then bMirroring = True   ' rules out AlphaBlend usage
        End If
        bStretching = Not (Abs(destWidth) = srcWidth) ' rules out AlphaBlend usage if Bilinear interpolation requested
        If destHeight = 0& Then
            destHeight = m_Height
        Else
            If destHeight < 0& Then bMirroring = True ' rules out AlphaBlend usage
        End If
        If Not bStretching Then bStretching = Not (Abs(destHeight) = srcHeight) ' rules out AlphaBlend usage if Bilinear interpolation requested
        
    If Me.isAlphaBlendFriendly Then ' Win98 or better with AlphaBlend enabled
        If Not bMirroring Then  ' Not mirroring, can use AlphaBlend
            If grayScale = AgsclNone Then
                If bStretching Then
                    bCanUseAlphaBlend = Not m_StretchQuality ' no Bilinear interpolation, can use AlphaBlend
                Else
                    bCanUseAlphaBlend = True
                End If
            End If
        End If
    End If
    
    If Me.isGDIplusEnabled = True And bCanUseAlphaBlend = False Then
        ' we will use GDI+ to render when higher quality interpolation is desired or system is not AlphaBlend friendly
        Dim cGDIp As New AcGDIPlus
        Dim Rg As Single, Gg As Single, Bg As Single
        Render = cGDIp.RenderGDIplus(Me, destinationDC, 0&, Opacity, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, m_StretchQuality, grayScale, m_GDItoken)

    Else
        
        
        If m_hDC = 0& Then  ' do we have a DC to select our image into?
            tDC = GetDC(0&) ' if not create one, if ManageOwnDC=True, we will have one
            m_hDC = CreateCompatibleDC(tDC)
            ReleaseDC 0&, tDC
            hOldImage = SelectObject(m_hDC, m_Handle)
        Else
            ' we have a DC, but is the image selected into it?
            If m_prevObj = 0& Then hOldImage = SelectObject(m_hDC, m_Handle)
        End If
        
        If bCanUseAlphaBlend = False Then
            ' Ruled out use of AlphaBlend (preferred when GDI+ isn't available)
            ' Win95/NT4 - can't use AlphaBlend
            ' Mirroring or high quality interpolation stretching - can't use AlphaBlend
            
            ' doing it completely manually
            If bMirroring Then MirrorDIB SrcX, SrcY, srcWidth, srcHeight, destWidth, destHeight, aMirrorBytes()
            If bStretching Then
                If pvResize(destinationDC, aResizedBytes(), aMirrorBytes(), Nothing, SrcX, SrcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight) = False Then Exit Function
            End If
            ' use custom blending routine
            Render = Win9xBlend(destinationDC, aResizedBytes(), SrcX, SrcY, destX, destY, destWidth, destHeight, (255& * Opacity) \ 100&, destHostDIB, grayScale)
        
        Else ' we can use AlphaBlend
        
            ' Stretch_Halftone not compatible with win9x
            If SetHalfTone Then
                If Me.isAlphaBlendFriendly Then lStretchMode = SetStretchBltMode(destinationDC, STRETCH_HALFTONE)
            End If
            
            ' calculate the opacity required & add it to the BlendFunction variable
            lBlendFunc = AC_SRC_OVER Or (((255& * Opacity) \ 100&) * &H10000)
            ' if the image has transparency, then we add the AC_SRC_ALPHA flag too
            If m_AlphaImage = True And Blend = True Then lBlendFunc = lBlendFunc Or (AC_SRC_ALPHA * &H1000000)
            Render = Not (AlphaBlend(destinationDC, destX, destY, destWidth, destHeight, m_hDC, SrcX, SrcY, srcWidth, srcHeight, lBlendFunc) = 0&)
            
            If SetHalfTone Then
                If Me.isAlphaBlendFriendly Then SetStretchBltMode destinationDC, lStretchMode
            End If
            
        End If
        
        ' remove the image from the DC if necessary
        If Not hOldImage = 0& Then SelectObject m_hDC, hOldImage
        If Not tDC = 0& Then    ' if we created a DC, let's destroy it now
            DeleteDC m_hDC
            m_hDC = 0&
        End If
        
    End If
    
End Function

Private Function LoadPictureEx(FileHandle As Long, FileName As String, aStream() As Byte, cx As Long, cy As Long, streamOffset As Long, streamLength As Long, SaveFormat As Boolean, bitDepth As Long) As Boolean
    
    ' PURPOSE: Marshal passed file/array to image classes for conversion to 32bpp image
    ' For parameter information, see LoadPicture_File & LoadPicture_Stream
    
    Me.DestroyDIB
    
    ' various image parsers, in order of precedence
    ' All 4 recognize transparency
    Dim cPNG As AcPNGparser  ' very fast to abort if not a PNG file
    
    Dim bReturn As Boolean  ' function return value
    Dim rtnRead As Long
    
    ' validate passed desired icon sizes
    If cx < 0& Then cx = 0&
    If cy < 0& Then cy = 0&
    If bitDepth < 0& Then
        bitDepth = 32
    ElseIf bitDepth > 32 Then
        bitDepth = 32
    End If
    
    Set cPNG = New AcPNGparser   ' see if image is a PNG; aborts quickly if not
    bReturn = cPNG.LoadFile(FileHandle, FileName, Me, m_GDItoken)
        'If bReturn = True Then Close #fileNum         ' close the file
    If Err Then MsgBox Err.Description
    Set cPNG = Nothing
    If Not bReturn Then
        If Not FileHandle = 0& Then
            streamOffset = 0&
            streamLength = GetFileSize(FileHandle, 0&)
            'streamLength = LOF(fileNum) ' cache length of file
            ReDim aStream(streamOffset To streamLength - 1&)
            'Get #fileNum, , aStream()   ' populate our stream with the file contents
            'Close #fileNum
            SetFilePointer FileHandle, 0&, 0&, 0&
            ReadFile FileHandle, aStream(streamOffset), streamLength, rtnRead, ByVal 0&
        End If
        'Set cGIF = New AcGIFparser ' what about a GIF; aborts quickly if not
        'bReturn = cGIF.LoadStream(aStream(), Me, streamOffset, streamLength)
        'Set cGIF = Nothing
        If Not bReturn Then
            'Set cICO = New AcICOparser   ' will process Vista PNG icon if needed
            'bReturn = cICO.LoadStream(aStream(), cx, cy, Me, streamOffset, streamLength, bitDepth, m_GDItoken)
            'Set cICO = Nothing
            If Not bReturn Then ' check for bmp, emf, wmf & jpg << last chance
                'Set cBMP = New AcBMPparser
                'bReturn = cBMP.LoadStream(aStream(), Me, streamOffset, streamLength)
                'Set cBMP = Nothing
            End If
        End If
    End If
    If Not m_Handle = 0 Then
        If SaveFormat = True Then ' we will cache the original bytes
            If AiparseIsArrayEmpty(VarPtrArray(aStream)) = 0& And Not FileHandle = 0& Then
                ' we loaded the image from the file and not a stream (PNG), need to get stream
                'fileNum = FreeFile()
                'Open FileName For Binary Access Read As #fileNum
                If streamLength = 0& Then streamLength = GetFileSize(FileHandle, 0&)
                ReDim m_ImageByteCache(0 To streamLength - 1)
                SetFilePointer FileHandle, 0&, 0&, 0&
                ReadFile FileHandle, m_ImageByteCache(0), streamLength, rtnRead, ByVal 0&
                'Get #fileNum, 1, m_ImageByteCache
                'Close #fileNum
            Else
                m_ImageByteCache() = aStream()
            End If
        End If
        LoadPictureEx = True
    End If

End Function

Private Function pvResize(ByVal destDC As Long, rSizedBytes() As Byte, rMirror() As Byte, Optional tHost As Ac32bppDIB, Optional ByVal SrcX As Long, Optional ByVal scrY As Long, Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, Optional ByVal destX As Long, Optional ByVal destY As Long, Optional destWidth As Long, Optional destHeight As Long) As Boolean
                            
    ' Function resizes an alpha image, maintaining premultiplied pixels & alpha values
    ' Code originally by Carles P.V. but significantly modified for this project.
    
    ' Parameters:
    ' destDC :: DC being rendered to, may be null
    ' rSizedbytes() : array to hold resized alpha section; not used if tHost is not Nothing
    ' tHost : when resizing to another DIB class, the destination DIB class
    ' srcX,Y : the coordinates of the source image to start resizing from
    ' srcWidth,srcHeight : the width/height of the source image to resize from
    ' destX,Y : the coordinates of the destination image to resize to
    ' destWidth,destHeight : the width/height of the destination image to resize to

    If srcWidth = 0& Then srcWidth = m_Width
    If srcHeight = 0& Then srcHeight = m_Height

    Dim aNewBits() As Byte, dSA As SafeArray   ' new size, overlay of DIB pointer
    Dim aOldBits() As Byte, tSA As SafeArray   ' old size, overlay of DIB pointer
    
    Dim xLUdbl() As Double                     ' look up table (LUT)
    Dim xRatio As Double, yRatio As Double     ' scaled ratios
    Dim srcPixel As Long, dstPixel As Long     ' source/destination pixel locations
    Dim lCol As Long, lRow As Long             ' loop variables
    
    Dim newWidth As Long, newHeight As Long
    
    ' following used with BiLinear scaling
    Dim fX As Double, fY As Double
    Dim tgtX As Long, tgtY As Long
    Dim edgeOffsetX As Long, edgeOffsetY As Long
    Dim iX As Long, iY As Long
    Dim R As Double, G As Double, B As Double, A As Double
    Dim scalerX As Double, scalerY As Double
    
    ' fill in opitonal parameters
    If Not tHost Is Nothing Then
        newWidth = tHost.Width
        newHeight = tHost.Height
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight
        If newWidth > tHost.Width Then newWidth = tHost.Width
        If newHeight > tHost.Height Then newHeight = tHost.Height
    Else
        newWidth = Abs(destWidth)
        newHeight = Abs(destHeight)
        
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight

        ' safety checks, recalculation of bounding destination size
        ' if not done, we could very easily access unallocated memory.
        If destX < 0 Then   ' negative DC offset
            newWidth = newWidth + destX    ' reduce width to process
            destX = -destX                 ' used to offset LUT; adjust so not processing bytes not used
        Else
            destX = 0&                        ' fits within destination bitmap; no offsetting needed
        End If
        
        ' now to check the vertical
        If destY < 0& Then  ' negative DC offset
            newHeight = newHeight + destY
            destY = 0&
        Else
            destY = 0&
        End If
        
    End If
    If newHeight < 1& Or newWidth < 1& Then Exit Function
    
    With dSA    ' overlay destination array onto the passed Byte() array
        .cbElements = 1
        .cDims = 2
        .rgSABound(0).cElements = newHeight
        .rgSABound(1).cElements = newWidth * 4&
        If tHost Is Nothing Then
            ReDim rSizedBytes(0& To .rgSABound(1).cElements - 1&, 0& To newHeight - 1&)
            .pvData = VarPtr(rSizedBytes(0&, 0&))
        Else
            .pvData = tHost.BitsPointer ' called by CopyImageTo & Resize routines
        End If
    End With
    
    With tSA    ' overlay source array onto our DIB
        .cbElements = 1
        .cDims = 2
        If AiparseIsArrayEmpty(VarPtrArray(rMirror)) = 0& Then
            .pvData = m_Pointer                     ' source is our DIB
            .rgSABound(0).cElements = m_Height
            .rgSABound(1).cElements = m_Width * 4&
        Else
            .pvData = VarPtr(rMirror(0&, 0&))           ' source is the mirrored DIB; clipped as needed
            .rgSABound(0).cElements = UBound(rMirror, 2) + 1&
            .rgSABound(1).cElements = UBound(rMirror, 1) + 1&
        End If
    End With
    
    CopyMemory ByVal VarPtrArray(aNewBits), VarPtr(dSA), 4&
    CopyMemory ByVal VarPtrArray(aOldBits), VarPtr(tSA), 4&
    
    On Error GoTo eh
    
    scrY = UBound(aOldBits, 2) - srcHeight - scrY + 1&     ' adjust Y position in source for bottom up DIBs
    If (m_StretchQuality = False) Then
        
        ' Scaling LUT, cache actual X position of DIB in relation to scaled X
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        ReDim xLUdbl(0 To newWidth - 1&)
        For lCol = 0 To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUdbl(lCol) = Int(((lCol + destX) * xRatio) + SrcX) * 4&
        Next
        
        ' nearest neighbor algorithm
        For lRow = newHeight - 1& To 0& Step -1&
            '^ current scanline for the scaled image
            ' offset destY is used for negative coordinates
            srcPixel = Int((lRow + destY) * yRatio) + scrY  ' recalcualted once per scanline
            ' current scanline for the scaled image
            dstPixel = 0&
            For lCol = 0& To newWidth - 1&
                ' copy into resized array the nearest raw/actual pixel
                CopyMemory aNewBits(dstPixel, lRow), aOldBits(Int(xLUdbl(lCol)), srcPixel), 4&
                dstPixel = dstPixel + 4&
            Next lCol
        Next lRow
        
    Else
        
        ' BiLinear interoplation, up to 4 source pixels (16 bytes) are blended for each destination pixel (4 bytes)
 
        ReDim xLUd(0 To newWidth - 1&)   ' work with doubles, we need the decimal portions
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        For lCol = 0& To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUd(lCol) = (((lCol + destX) * xRatio) + SrcX)
        Next
        srcWidth = srcWidth - 1& ' subtract now vs subtracting in loop below
        For lRow = newHeight - 1& To 0& Step -1&
            
            fY = (lRow + destY) * yRatio + scrY ' get the scaled source row
            tgtY = Int(fY)              ' get whole number of double
            If tgtY = 0& Then         ' for top down images, test for last row vs 0
                ' last row of source image, will use only this row
                edgeOffsetY = 0&
                fY = 0#
            Else
                ' will use this row & next row for blending
                edgeOffsetY = 1&
                fY = Abs(1# - (fY - tgtY)) ' for top down images, use fY = fY-tgtY
                ' ^ Y coordinate fraction; pct of next vertical pixel that is used
            End If
            
            For lCol = 0& To newWidth - 1&
                
                R = 0#: G = 0#: B = 0#: A = 0#
                
                tgtX = Int(xLUd(lCol))  ' coordinate rounded down to whole number
                If tgtX = srcWidth Then
                    ' at far edge of source image, will use only this pixel for blending
                    edgeOffsetX = 0&
                    fX = 0#
                Else
                    ' will use this pixel and next pixel for blending
                    edgeOffsetX = 1&
                    fX = xLUd(lCol) - tgtX
                    ' ^ X coordinate fraction, pct of next horizontal pixel that is used
                End If
                
                For iY = 0& To edgeOffsetY
    
                    scalerY = Abs(iY - fY)   ' percentage of current row's pixel to blend
                    If Not scalerY = 1# Then  ' else result will be zero
                    
                        For iX = 0& To edgeOffsetX
                            scalerX = Abs(fX - iX)   ' percentage of current column's pixel to blend
                            
                            If Not scalerX = 1# Then  ' else result will be zero
                                scalerX = (1# - scalerX) * (1# - scalerY) ' combine percentages
    
                                ' Build the blended RGB values, for top down images Add iY vs subtracting
                                srcPixel = (tgtX + iX) * 4&
                                B = B + aOldBits(srcPixel, tgtY - iY) * scalerX
                                G = G + aOldBits(srcPixel + 1&, tgtY - iY) * scalerX
                                R = R + aOldBits(srcPixel + 2&, tgtY - iY) * scalerX
                                A = A + aOldBits(srcPixel + 3&, tgtY - iY) * scalerX
    
                            End If
                        Next
                    End If
                Next
                iX = lCol * 4&
                ' update destination with adjusted pixel
                aNewBits(iX, lRow) = Int(B)
                aNewBits(iX + 1&, lRow) = Int(G)
                aNewBits(iX + 2&, lRow) = Int(R)
                aNewBits(iX + 3&, lRow) = Int(A)
            Next
        Next
    End If
    
    CopyMemory ByVal VarPtrArray(aOldBits), 0&, 4&
    CopyMemory ByVal VarPtrArray(aNewBits), 0&, 4&
    
    ' the passed destWidth,destHeight params are used when rendering; we are just sizing now
    destWidth = newWidth ' the parameter is ByRef, update it now
    destHeight = newHeight ' the parameter is ByRef, update it now
    Erase rMirror()
    pvResize = True
eh:
    If Err Then
        Err.Clear   ' troubleshooting only, should be removed before compiling to final app
        Stop
        Resume
    End If
End Function

Private Function Win9xBlend(ByVal destinationDC As Long, aResizedBytes() As Byte, ByVal SrcX As Long, ByVal SrcY As Long, ByVal destX As Long, ByVal destY As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal GlobalAlpha As Long, tHost As Ac32bppDIB, ByVal grayScale As AeGrayScaleFormulas) As Boolean

    ' Function manually blends an alpha bitmap to a target DC (used primarily for Win95 & NT4)
    ' and when rotating, high quality stretching & mirroring when GDI+ is not available
    
    ' Parameters identify the destination more than anything else. The source was already pre-processed if needed
    
    ' destinationDC :: DC to blend to
    ' aResizedBytes() :: array of bytes sized to target destination blend area.
    '   if array is null, then the destination size is same size as our DIB's image
    ' srcX,Y :: the position on source to begin blending; only used if aResizedBytes() is null
    ' destX,Y :: the position on destination where blending starts
    ' destWidth,Height :: the amount of columns/rows to blend
    ' globalAlpha :: the AlphaBlend global alpha value: between 0 and 255
    ' tHost (Optional) :: blending will occur DIB to DIB vs DIB to DC
    
    ' Special note. Having problems rendering 32bpp DIBs to WinME; artifacts are being rendered
    ' from the alpha channel. Therefore, to completely eliminate this problem (hopefully), the
    ' destination array will be 24bpp vs 32bpp which will then be updated and rendered onto the
    ' destination DC.  32bpp would be easier, but oh well.

    Dim srcBytes() As Byte, srcSA As SafeArray
    Dim dstBytes() As Byte, dstSA As SafeArray
    Dim srcCol As Long, srcRow As Long
    Dim srcAlpha As Long, dstAlpha As Long
    Dim Y As Long, X As Long
    Dim sX As Long, sY As Long
    Dim dX As Long, dY As Long
    Dim dDC As Long, tDC As Long, hOldBmp As Long, hDib As Long
    Dim mirrorOffsetX As Long, mirrorOffsetY As Long
    Dim Rg As Single, Gg As Single, Bg As Single
    
    Dim BMPI As BITMAPINFO
    
    ' The following is just a wee bit confusing.
    ' Our source can be 2 different objects:
    '   1) Our host DIB if the rendering results with no stretching
    '   2) The passed aResizedBytes() array if stretching was involved
    ' Likewise, the destination can be 2 different objects
    '   1) A passed DC handle
    '   2) Another DIB class if tHost is passed
    ' So to use a common base for all possibilities, we use SafeArrays
    
    ' ensure we can update the target DC and also get the bitmap dimensions
    If destWidth < 0& Then           ' handle horizontal mirroring
        destWidth = -destWidth
        mirrorOffsetX = -4&          ' flag to be used later
    End If
    If destHeight < 0& Then          ' handle vertical mirroring
        destHeight = -destHeight
        mirrorOffsetY = 1&           ' flag to be used later, row read top down
    Else
        mirrorOffsetY = -1           ' no mirroring rows are read bottom up
    End If
    If AiparseIsArrayEmpty(VarPtrArray(aResizedBytes)) = 0& Then
         srcSA.pvData = m_Pointer
         ' need to tweak for negative offsets
         If destX < 0& Then
             SrcX = SrcX - destX ' less area that needs to be rendered
             destWidth = destWidth + destX ' adjust amount of destination we copy
             destX = 0&                    ' set destination offset to zero
         End If
         If destY < 0& Then
             SrcY = SrcY - destY ' less area that needs to be rendered
             destHeight = destHeight + destY ' adjust amount of destination we copy
             destY = 0&                      ' set destination offset to zero
         End If
         SrcY = m_Height - SrcY  ' set DIB offset for 1st row to be blended
         
     Else ' source is the resized array
         srcSA.pvData = VarPtr(aResizedBytes(0, 0))
         If destX < 0& Then destX = 0& ' when negative target coords are used, we set to
         If destY < 0& Then destY = 0& ' zero to match the resized array (zero-based)
         SrcX = 0&: SrcY = destHeight
    End If
     
     If Not tHost Is Nothing Then
        ' need to ensure we won't be copying memory to unallocated memory
        If destWidth > tHost.Width Then destWidth = tHost.Width
        If destHeight > tHost.Height Then destHeight = tHost.Height
     End If
     If destWidth < 1& Or destHeight < 1& Then Exit Function  ' nothing to Blt, passed bad params
    
    If tHost Is Nothing Then
        ' we need the contents of the target DC, describe its bitmap
        With BMPI.bmiHeader
            .biSize = 40
            .biHeight = destHeight
            .biWidth = destWidth
            .biPlanes = 1
            .biBitCount = 24    ' see Special note above
        End With
        dDC = GetDC(0&)
        tDC = CreateCompatibleDC(dDC)
        If Not tDC = 0& Then
            hDib = CreateDIBSection(dDC, BMPI, 0&, dstSA.pvData, 0&, 0&)
        End If
        ReleaseDC 0&, dDC
        If tDC = 0& Or hDib = 0& Then Exit Function
        hOldBmp = SelectObject(tDC, hDib)
        BitBlt tDC, 0&, 0&, destWidth, destHeight, destinationDC, destX, destY, vbSrcCopy
        SelectObject tDC, hOldBmp
        DeleteDC tDC
        
        With dstSA
            .cbElements = 1
            .cDims = 2
            .rgSABound(0).cElements = destHeight
            .rgSABound(1).cElements = AiparseByteAlignOnWord(24, destWidth)
        End With
    
    Else
        ' DIB class to DIB class blending
        With dstSA
            .cbElements = 1
            .cDims = 2
            .pvData = tHost.BitsPointer
            .rgSABound(0).cElements = tHost.Height
            .rgSABound(0).lLbound = -(tHost.Height - destY - destHeight)
            .rgSABound(1).cElements = tHost.scanWidth
            .rgSABound(1).lLbound = -destX * 4
        End With
    End If
                    
    CopyMemory ByVal VarPtrArray(dstBytes), VarPtr(dstSA), 4&
    
    With srcSA  ' overlay onto our DIB
        .cbElements = 1
        .cDims = 2
        If .pvData = m_Pointer Then ' using DIB as source
            .rgSABound(0).cElements = m_Height
            .rgSABound(1).cElements = m_Width * 4&
        Else                        ' using resized array as source
            .rgSABound(0).cElements = UBound(aResizedBytes, 2) + 1
            .rgSABound(1).cElements = UBound(aResizedBytes, 1) + 1
        End If
    End With
    CopyMemory ByVal VarPtrArray(srcBytes), VarPtr(srcSA), 4&
                    
    On Error Resume Next ' expected errors? using a corrupted pre-multiplied image possibly, but will just draw wrong
    
    If mirrorOffsetY = 1& Then SrcY = SrcY - 1  ' handle horizontal mirroring
    If mirrorOffsetX = -4& Then                 ' handle horizontal mirroring
        SrcX = (SrcX + destWidth - 1&) * 4&     ' last pixel col will be destination's first
    Else
        mirrorOffsetX = 4&
        SrcX = SrcX * 4&
    End If
    
    ' this loop is broken into several different loops to enhance speed when less options are used
    ' 1. separate loop when rendering DIB to DC or DIB to DIB
    ' 2. separate loop when using global alpha of 255 or using less than 255
    ' 3. separate loop when grayscaling and not grayscaling
    
    For srcRow = 0& To destHeight - 1&
    
        ' offset our DIB row as we go
        SrcY = SrcY + mirrorOffsetY     ' adjusted source row, handle vertical mirroring
        sX = SrcX                       ' 1st column of source row, includes horizontal mirroring
        dY = destHeight - srcRow - 1    ' next row for destination image
        dX = 0&                         ' 1st column of destination row
        
        If tHost Is Nothing Then        ' DIB to DC rendering
            
            If grayScale = AgsclNone Then
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             CopyMemory dstBytes(dX, dY), srcBytes(sX, SrcY), 3&
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                         End If
                        dX = dX + 3&
                        sX = sX + mirrorOffsetX
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                         End If
                         dX = dX + 3&
                         sX = sX + mirrorOffsetX
                     Next
                End If
            Else    ' gray scaling
                Call AiparseGrayScaleRatios(grayScale, Rg, Gg, Bg)
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             dstBytes(dX, dY) = (((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)) * 255&) \ 255
                             dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                            
                            dstBytes(dX, dY) = (((dstBytes(dX, dY) * Bg) + (dstBytes(dX + 1&, dY) * Gg) + (dstBytes(dX + 2&, dY) * Rg)) * 255&) \ 255
                            dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                            dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                         End If
                        dX = dX + 3&
                        sX = sX + mirrorOffsetX
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                         
                            dstBytes(dX, dY) = (((dstBytes(dX, dY) * Bg) + (dstBytes(dX + 1&, dY) * Gg) + (dstBytes(dX + 2&, dY) * Rg)) * 255&) \ 255
                            dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                            dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                         End If
                         dX = dX + 3&
                         sX = sX + mirrorOffsetX
                     Next
                End If
            End If
    
        Else    ' DIB class to DIB class rendering
        
            If grayScale = AgsclNone Then
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             CopyMemory dstBytes(dX, dY), srcBytes(sX, SrcY), 4&
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                            dstBytes(dX + 3&, dY) = (dstAlpha * dstBytes(dX + 3&, dY)) \ &HFF + srcBytes(sX + 3&, SrcY)
                         End If
                        dX = dX + 4&
                        sX = sX + mirrorOffsetX
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 3&, dY) = (dstAlpha * (dstBytes(dX + 3&, dY)) + (srcBytes(sX + 3, SrcY) * GlobalAlpha)) \ &HFF
                         End If
                         dX = dX + 4&
                         sX = sX + mirrorOffsetX
                     Next
                End If
            Else
                Call AiparseGrayScaleRatios(grayScale, Rg, Gg, Bg)
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             dstBytes(dX, dY) = (((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)) * 255&) \ 255
                             dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 3&, dY) = srcBytes(sX + 3&, SrcY)
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                            dstBytes(dX + 3&, dY) = (dstAlpha * dstBytes(dX + 3&, dY)) \ &HFF + srcBytes(sX + 3&, SrcY)
                         
                            dstBytes(dX, dY) = (((dstBytes(dX, dY) * Bg) + (dstBytes(dX + 1&, dY) * Gg) + (dstBytes(dX + 2&, dY) * Rg)) * 255&) \ 255
                            dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                            dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                         End If
                        dX = dX + 4&
                        sX = sX + mirrorOffsetX
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 3&, dY) = (dstAlpha * (dstBytes(dX + 3&, dY)) + (srcBytes(sX + 3, SrcY) * GlobalAlpha)) \ &HFF
                         
                            dstBytes(dX, dY) = (((dstBytes(dX, dY) * Bg) + (dstBytes(dX + 1&, dY) * Gg) + (dstBytes(dX + 2&, dY) * Rg)) * 255&) \ 255
                            dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                            dstBytes(dX + 2&, dY) = dstBytes(X, Y)
                         End If
                         dX = dX + 4&
                         sX = sX + mirrorOffsetX
                     Next
                End If
            End If
        End If
    Next
    ' remove overlay
    CopyMemory ByVal VarPtrArray(srcBytes), 0&, 4&
    ' transfer results
    If tHost Is Nothing Then
        SetDIBitsToDevice destinationDC, destX, destY, destWidth, destHeight, 0&, 0&, 0&, destHeight, dstBytes(0, 0), BMPI, 0&
    End If
    CopyMemory ByVal VarPtrArray(dstBytes), 0&, 4&
    If Not hDib = 0& Then DeleteObject hDib
    
    Win9xBlend = True

End Function

Private Sub MirrorDIB(ByRef SrcX As Long, ByRef SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long, ByRef mirrorBytes() As Byte, Optional ByRef tHost As Ac32bppDIB)

    ' through trial and error, the rule, for mirroring images, appears to be valid for
    ' Windows drawing routines are:  Flip first, then stretch and/or rotate, then clip as needed
    
    ' For manual rendering, this has the potential of really complex offsets whether or not
    ' the pixels are read left to right, right to left, top to bottom or vice versa. Not
    ' to mention rotation and/or scaling and offsets whether or not image is to be clipped
    ' because it is being rendered at negative X,Y coordinates or the image is clipped
    ' because its scaled size is too big for the destination bitmap area. So, we will
    ' not precalculate all the possible combinations of offsets in the manual rendering
    ' routines, rather we will flip first, then allow rendering routines to process as normal.
    
    ' srcX,Y :: position in source image where mirroring begins
    ' srcWidth,Height :: amount of source image that will be mirrored
    ' newWidth,Height :: size of destination mirrored image
    ' mirrorBytes() :: byte array to hold mirrored image
    ' tHost :: called by CopyImageTo & Resize when mirroring
    
    ' is image being mirrored?
    If newWidth > 0& And newHeight > 0& Then Exit Sub
    
    Dim tSA As SafeArray, srcBytes() As Byte
    Dim dSA As SafeArray, dstBytes() As Byte
    Dim X As Long, Y As Long, yOffset As Long, xOffset As Long
    
    With tSA    ' overlay array on our source DIB
        .cbElements = 1
        .cDims = 2
        .pvData = m_Pointer
        .rgSABound(0).cElements = m_Height
        .rgSABound(1).cElements = m_Width * 4&
    End With
    
    With dSA    ' overlay array on our destination object
        .cbElements = 1
        .cDims = 2
        If tHost Is Nothing Then    ' destination is mirrorBytes
            ReDim mirrorBytes(0& To srcWidth * 4& - 1&, 0& To srcHeight - 1&)
            .pvData = VarPtr(mirrorBytes(0&, 0&))
        Else                        ' destination is passed DIB class
            .pvData = tHost.BitsPointer ' called by CopyImageTo & Resize routines
        End If
        .rgSABound(0).cElements = srcHeight
        .rgSABound(1).cElements = srcWidth * 4&
    End With
    
    CopyMemory ByVal VarPtrArray(srcBytes), VarPtr(tSA), 4&
    CopyMemory ByVal VarPtrArray(dstBytes), VarPtr(dSA), 4&
    
    If newHeight < 0& Then
        If newWidth > 0& Then   ' flipping vertically only, faster/easier
            xOffset = srcWidth * 4& ' number of bytes to flip at once
            Y = SrcY    ' when flipping, our Y is adjusted from the bottom
            X = SrcX * 4&   ' starting X position in source
            For yOffset = srcHeight - 1& To 0& Step -1&
                CopyMemory dstBytes(0&, yOffset), srcBytes(X, Y), xOffset
                Y = Y + 1&   ' move source Y to next row
            Next
        Else   ' flipping both vertically/horizontally
            Y = SrcY ' when flipping vertically, adjust from the bottom
            For yOffset = srcHeight - 1& To 0& Step -1&
                X = (m_Width - SrcX) * 4& - 4&  ' X adjusted from right when flipping
                For xOffset = 0 To srcWidth * 4 - 4& Step 4&
                    CopyMemory dstBytes(xOffset, yOffset), srcBytes(X, Y), 4&
                    X = X - 4&  ' move source X to next pixel
                Next
                Y = Y + 1&   ' move source Y to next row
            Next
        End If
    
    Else  ' flipping horizontally only, same comments as above
        Y = m_Height - SrcY - 1& ' not fipping vertically, so read from bottom up
        For yOffset = srcHeight - 1& To 0& Step -1&
            X = (m_Width - SrcX) * 4& - 4&
            For xOffset = 0 To srcWidth * 4 - 4& Step 4&
                CopyMemory dstBytes(xOffset, yOffset), srcBytes(X, Y), 4&
                X = X - 4&
            Next
            Y = Y - 1&
        Next
    End If
    CopyMemory ByVal VarPtrArray(srcBytes), 0&, 4&
    CopyMemory ByVal VarPtrArray(dstBytes), 0&, 4&
    
    ' any rendering routines will now use the mirrored byte array
    ' so we ensure the bounds parameters match the mirrored array bounds
    SrcX = 0&
    SrcY = 0&
    newWidth = Abs(newWidth)
    newHeight = Abs(newHeight)
    
End Sub

Private Sub Class_Initialize()

    ' Determine operating system for compatibility of 32bpp images
    ' http://vbnet.mvps.org/code/helpers/iswinversion.htm
    ' http://msdn2.microsoft.com/en-gb/library/ms724834.aspx
    
   Dim osType As OSVERSIONINFOEX
   Const VER_PLATFORM_WIN32_WINDOWS As Long = 1

   ' Retrieve version data for OS.
   osType.dwOSVersionInfoSize = Len(osType)
   If GetVersionEx(osType) = 0 Then
      ' The OSVERSIONINFOEX structure is only supported
      ' in NT4/SP6+ and NT5.x, so we're likely running
      ' on an earlier version of Windows. Revert structure
      ' size to OSVERSIONINFO and try again.
      osType.dwOSVersionInfoSize = Len(osType) - 8
      Call GetVersionEx(osType)
   End If
   
   ' How the m_osCAP variable works and is used througout this class
   ' Value contains 1, then AlphaBlend enabled & used when needed. Not enabled on Win9x unless overridden via isAlphaBlendFriendly
   ' Value contains 2, then GDI+ enabled & used when needed (set in isGDIplusEnabled)
   ' Value contains 4, then zLib enabled & can be used to create/read PNGs (set in isZlibEnabled). Any O/S :: no longer used; tested as needed
   ' Value contains 8, then Win98+ or Win2K+: AlphaBlend capable system else it isn't
   ' Value contains 16, then a Win98 or WinME system
   ' Value contains 32, then NT4 w/less than SP6 or Win95. Otherwise system is GDI+ capable else it isn't
   ' Note: when m_osCap contains 17, then AlphaBlend has been overridden by user
   
   
    If osType.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then

        If osType.dwMinorVersion = 0 Then ' Win95; can't use AlphaBlend nor GDI+
            m_osCAP = 32

        Else ' flag as Alphablend disabled, but capable & is Win98/ME
            m_osCAP = 8 Or 16
        End If

    Else

        If osType.dwMajorVersion > 4 Then ' if Win2K or better
            m_osCAP = 1 Or 8    ' flag as AlphaBlend enabled (Win2K or better) and capable

        Else ' WinNT4. If SP6 or better than GDI+ capable else not. Regardless, not AlphaBlend capable
            If osType.wServicePackMajor < 6 Then m_osCAP = 32
        End If

    End If

    Me.isGDIplusEnabled = True  ' attempt to start GDI+, test system capability
    If Me.isGDIplusEnabled Then Me.HighQualityInterpolation = True


End Sub

Private Sub Class_Terminate()
    DestroyDIB ' simply clean up
End Sub
