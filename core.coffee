class Core

  # Member
  debug: false

  fso: new ActiveXObject "Scripting.FileSystemObject"

  # Static method# {{{
  @isFile: (path) -># {{{
    @fso.FileExists path# }}}

  @isDirectory: (path) -># {{{
    @fso.FolderExists path# }}}

  @joinPath: (paths...) -># {{{
    root = paths.shift()
    for path in paths
      root = @fso.BuildPath root, path
    return root# }}}

  @copyFile: (src, dst, overwrite = false) -># {{{
    @fso.CopyFile src, dst, overwrite# }}}

  @copyDirectory: (src, dst, overwrite = false) -># {{{
    @fso.CopyFolder src, dst, overwrite# }}}

  @copy: (src, dst, overwrite = false) -># {{{
    if @isFile.call new FSO src
      if @isDirectory.call new FSO dst
        @fso.CopyFile src, @joinPath(dst, src), overwrite
      else
        @fso.CopyFile src, dst, overwrite
    else
      @copyDirectory src, dst, overwrite# }}}

  @moveFile: (src, dst, overwrite = false) -># {{{
    if @isFile dst and overwrite
      @delete.call new FSO dst
    else
      @fso.MoveFile src, dst# }}}

  @moveDirectory: (src, dst, overwrite = false) -># {{{
    if @isDirectory dst and overwrite
      @fso.MoveFolder "#{src}\*", dst
    else if not @isDirectory dst
      @fso.MoveFolder src, dst# }}}

  @walk: (path, filter) -># {{{
    @dp "path = [#{path}]"
    ret = []
    do f = (path) =>
      dir = @fso.GetFolder path
      # directories
      e = new Enumerator dir.SubFolders
      until e.atEnd()
        f new String e.item()
        e.moveNext()
      # files
      e = new Enumerator dir.Files
      until e.atEnd()
        ret.push new String e.item()
        e.moveNext()
    return ret# }}}
# }}}

  # Constructor# {{{
  constructor: (@path = path) -># {{{
    super()
    if @fso.FileExists @path
      @file = @fso.GetFile @path
    else
      @dir = @fso.GetFolder @path
      @isroot = @dir.IsRootFolder# }}}
# }}}

  # Instance method# {{{
  p: (msg) -># {{{
    WScript.Echo msg# }}}

  dp: (msg) -># {{{
    WScript.Echo msg if @debug# }}}

  filterArray: (array, filter) -># {{{
    @dp "filter = [#{filter}]"
    ret = []
    for a in array
      if a.match eval "/#{filter}/"
        @dp a
        ret.push a
    return ret# }}}

  createDirectory: (path) -># {{{
    do f = (path) =>
      if @isDirectory.call new Core @fso.GetParentFolderName path
        @fso.CreateFolder path
      else
        f @fso.GetParentFolderName path# }}}

  exists: -># {{{
    @fso.FileExists @path
    @fso.FolderExists(@path)# }}}

  copyTo: (destination, overwrite = false) -># {{{
    @file.Copy destination, overwrite# }}}

  delete: (force = false) -># {{{
    @file.Delete()
    @dir.Delete()# }}}

  moveTo: (destination) -># {{{
    @file.Move destination# }}}

  getLastAccessed: -># {{{
    date = @file.DataLastAccessed
    return null unless date
    new Date(date)# }}}

  getCreatedDate: -># {{{
    date = @file.DateCreated
    return null unless date
    new Date(date)# }}}

  getLastModified: -># {{{
    date = @file.DateLastModified
    return null unless date
    new Date(date)# }}}

  getSize: -># {{{
    @file.Size# }}}

  getBaseName: -># {{{
    @fso.GetBaseName @path# }}}

  setBaseName: (str) -># {{{
    @file.Name = "#{str}.#{@getExtensionName}"# }}}

  getExtensionName: -># {{{
    @fso.GetExtensionName @path# }}}

  setExtensionName: (str) -># {{{
    @file.Name = "#{@getBaseName}.#{str}"# }}}

  getName: -># {{{
    @file.Name# }}}

  getFullName: -># {{{
    @file.Path + String# }}}

  setName: (str) -># {{{
    @file.Name = str# }}}

  getParentDirectory: -># {{{
    @file.ParentFolder.Path# }}}

  read: -># {{{
    # some characters are broken in 'iso-8859-1'.
    illegalChars =
      0x20ac: 0x80, 0x81  : 0x81, 0x201a: 0x82, 0x192 : 0x83, 0x201e: 0x84,
      0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x2c6 : 0x88, 0x2030: 0x89,
      0x160 : 0x8a, 0x2039: 0x8b, 0x152 : 0x8c, 0x8d  : 0x8d, 0x17d : 0x8e,
      0x8f  : 0x8f, 0x90  : 0x90, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93,
      0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x2dc : 0x98,
      0x2122: 0x99, 0x161 : 0x9a, 0x203a: 0x9b, 0x153 : 0x9c, 0x9d  : 0x9d,
      0x17e : 0x9e, 0x178 : 0x9f

    throw new Error("File not found: " + file) unless @exists
    return "" if @getSize == 0
    # read in binary
    try
      stream = new ActiveXObject "ADODB.Stream"
      stream.Type = 2  # adTypeText
      stream.Charset = 'iso-8859-1'
      stream.Open()
      stream.LoadFromFile file
      text = stream.readText()
      list = []
      for i in [0...text.length]
        v = text.charCodeAt i
        list.push(illegalChars[v] or v)
      return String.fromCharCode.apply null, list
    finally
      stream.Close() if stream# }}}

  write: (str, encode) -># {{{
    new Directory(@getParentDirectory()).createFile @getName() unless @exists()
    try
      stream = new ActiveXObject "stream.Stream"
      if encode
        stream.Charset = encode
      else
        stream.Charset = "UTF-8"
      stream.Open()
      stream.Type = 2
      stream.WriteText str
      stream.SaveToFile @path, 2
    finally
      stream.close() if stream# }}}

  getSubDirectories: -># {{{
    ret = new Array()
    sub = new Enumerator @dir.SubFolders
    until sub.atEnd()
      ret.push @fso.BuildPath @getFullName(), sub.item().Name
      sub.moveNext()
    return ret# }}}

  getSubFiles: -># {{{
    return [] unless @exists()
    ret = new Array()
    files = new Enumerator @dir.Files
    until files.atEnd()
      ret.push @fso.BuildPath @getFullName(), files.item().Name
      files.moveNext()
    return ret# }}}

  getParentDirectory: -># {{{
    @dir.ParentFolder# }}}

  getName: -># {{{
    @dir.Name# }}}

  setName: (str) -># {{{
    @dir.Name = str# }}}

  getfullname: -># {{{
    @dir.path + string# }}}

  copyTo: (destination, overwrite = false) -># {{{
    @dir.Copy destination, overwrite# }}}

  moveTo: (destination) -># {{{
    @dir.Move destination# }}}

  createFile: (fileName, overwrite = false, encoding) -># {{{
    if encoding is "UTF-8"
      encoding = true
    else if encoding is "Shift_JIS"
      encoding = false
    else
      encoding = false
    @dir.CreateTextFile fileName, overwrite, encoding
    new File(@getFullName() + fileName)# }}}

  getLastAccessed: -># {{{
    date = @dir.DataLastAccessed
    return null  unless date
    new Date(date)# }}}

  getCreatedDate: -># {{{
    date = @dir.DateCreated
    return null  unless date
    new Date(date)# }}}

  getLastModified: -># {{{
    date = @dir.DateLastModified
    new Date(date)# }}}

  getSize: -># {{{
    @dir.Size# }}}
# }}}
