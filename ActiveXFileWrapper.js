File = function(path){
	this.Path = path ;
	this.fso = new ActiveXObject("Scripting.FileSystemObject")
	this.file = this.fso.GetFile(this.Path) ;
}

File.prototype = {
	exists : function(){
		if(!this.fso.FileExists(this.Path)){
			return true
		}
		return false
	} ,
	CopyTo : function(destination , overwrite){
		if(!overwrite) overwrite = false ; 
		this.file.Copy(destination , overwrite)
	} ,

	Delete : function(force){
		if(!force) force = false ;
		this.file.Delete() ;
	} ,

	Move : function(destination){
		this.file.Move(destination)
	} ,

	getLastAccessed : function(){
		var date = this.file.DataLastAccessed
		
		if(!date) return null ;
		return new Date(date) ;
	} ,
	
	getCreatedDate : function(){
		var date = this.file.DateCreated
		if(!date) return null ;
		return new Date(date) ;
	} ,
	getLastModified : function(){
		var date = this.file.DateLastModified
		return new Date(date) ;
	} ,
	getSize : function(){
		return this.file.Size
	} ,
	getBaseName : function(){
		return this.fso.GetBaseName(this.Path)
	} ,
	setBaseName : function(str){
		this.file.Name = str + "." + (this.fso.GetExtensionName(this.path))
	} ,
	getExtensionName : function(){
		return fso.GetExtensionName(this.Path) ;
	} ,
	setExtensionName : function(str){
		return this.file.Name = (this.fso.GetBaseName(this.Path)) + "." + str ;
	} ,
	getName : function(){
		return this.file.Name ;
	} ,
	setName : function(str){
		new this.file.Name = str ;
	} ,
	getParentDirectory : function(){
		return Directory( this.file.ParentFolder )
	} ,
	read : function(encode){
		var adodb = new ActiveXObject("ADODB.Stream") ;
		adodb.Charset = encode ;
		if(!encode) adodb.Charset = "Shift_JIS" ;
		adodb.Open()
		adodb.LoadFromFile( this.Path ) ;
		var text = adodb.ReadText() ;
		adodb.Close() ;
		return text
		
	} ,
	write : function(str){
		var stream =  this.fso.OpenTextFile(this.Path)
		var str = stream.Write(str) ;
		stream.Close() ;
		return str ;
	}
}
var Directory = function(path){
	if(path.substring( path.length -1 , path.length) != "\\")
		path += "\\"

	this.Path = path ;
	this.fso = new ActiveXObject("Scripting.FileSystemObject") ;
	this.checking() ;
	this.dir = this.fso.GetFolder( this.Path ) ;
	this.isroot = this.dir.IsRootFolder;
}
Directory.prototype = {
	checking : function(){
		if(!this.fso.FolderExists(this.Path)){
			throw new Exception.UnknownDirectoryException(this.Path) ;
		}
	} ,
	getSubDirectories : function(){
		var ret = [] ;
		var folders = new Enumerator(this.dir.SubFolders) ;
		
		for( ; !folders.atEnd() ; folders.moveNext() ){
			ret.push(new Directory( this.getFullName() + folders.item().Name)) ;
		}
		
		return ret ;
	} ,	
	getSubFiles : function(){
		this.checking() ;
		var ret = [] ;
		var files = new Enumerator(this.dir.files)
		
		for( ; !files.atEnd() ; files.moveNext() ){
			ret.push(new File(this.getFullName() + files.item().Name) ) ;
		}
		
		return ret ;
	} ,	
	getParentDirectory : function(){
		return Directory( this.dir.ParentFolder )
	} ,
	
	getName : function(){
		return this.dir.Name
	} ,
	setName : function(str){
		this.dir.Name = str ;
	} ,
	getFullName : function(){
		var str =  this.dir + ""
		if(str.substring( str.length -1 , str.length) != "\\")
			str += "\\"
		return str ;
	} ,
	copyTo : function( destination , overwrite){
		if(!overwrite) overwite = false ;
		this.dir.Copy(destination , overwrite) ;
	} ,
	Delete : function(){
		this.dir.Delete() ;
	} ,
	Move : function(destination){
		this.dir.Move(destination) ;
	} ,
	CreateFile : function(fileName , overwrite , encoding){
		if(!overwrite) overwrite = false ;
		if(encoding == "UTF-8") encoding = true
		else if(encoding == "Shift_JIS") encoding = false ;
		else encoding = false ;
		
		this.dir.CreateTextFile(fileName , overwrite , encoding) ;
		
	} ,
	CreateDirectory : function(folderName){
		this.fso.CreateFolder(this.getFullName() + folderName) ;
	} ,
	getLastAccessed : function(){
		var date = this.dir.DataLastAccessed
		
		if(!date) return null ;
		return new Date(date) ;
	} ,
	getCreatedDate : function(){
		var date = this.dir.DateCreated
		
		if(!date) return null ;
		return new Date(date) ;
	} ,	
	getLastModified : function(){
		var date = this.dir.DateLastModified
		
		return new Date(date) ;
	} ,
	getSize : function(){
		return this.dir.Size
	} ,
	toString : function(){ return "[object Directory]"}
} 