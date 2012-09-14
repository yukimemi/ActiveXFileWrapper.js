File = function(path){
			this.Path = path ;
}

File.prototype = {
	exists = function(){
		if(!new ActiveXObject("Scripting.FileSystemObject").FileExists(this.Path)){
			return true
		}
		return false
	} ,
	CopyTo : function(destination , overwrite){
		if(!overwrite) overwrite = false ; 
		new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.Copy(destination , overwrite)
	} ,

	Delete : function(force){
		if(!force) force = false ;
		new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.Delete() ;
	} ,

	Move : function(destination){
		new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.Move(destination)
	} ,

	getLastAccessed : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.DataLastAccessed
		
		if(!date) return null ;
		return new Date(date) ;
	} ,
	
	getCreatedDate : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.DateCreated
		if(!date) return null ;
		return new Date(date) ;
	} ,
	getLastModified : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject")
			.GetFile(this.Path)
			.DateLastModified
		
		
		return new Date(date) ;
	} ,
	getSize : function(){
		return new ActiveXObject("Scripting.FileSystemObject")
					.GetFile(this.Path)
					.Size
	} ,
	getBaseName : function(){
		return new ActiveXObject("Scripting.FileSystemObject")
					.GetBaseName(this.Path)
	} ,
	setBaseName : function(str){
		return new ActiveXObject("Scripting.FileSystemObject")
					.GetFile( this.path )
					.Name = str + "." + (new ActiveXObject("Scripting.FileSystemObject").GetExtensionName(this.path))
	} ,
	getExtensionName : function(){
		return new ActiveXObject("Scripting.FileSystemObject")
					.GetExtensionName(this.Path) ;
	} ,
	getParentDirectory : function(){
		return Directory( new ActiveXObject("Scripting.FileSystemObject").GetFile(this.Path).ParentFolder )
	} ,
	read : function(){
		return new ActiveXObject("Scripting.FileSystemObject")
					.OpenTextFile(this.Path)
					.ReadAll()
	} ,
	write : function(str){
		var stream =  new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(this.Path)
		var str = stream.Write(str) ;
		stream.Close() ;
		return str ;
	}
}
Directory = function(path){
	this.Path = path ;
}
Directory.prototype = {
	checking : function(){
		if(!new ActiveXObject("Scripting.FileSystemObject").FolderExists(this.Path)){
			throw new API.Exception.UnknownDirectoryException(this.Path) ;
		}
	} ,
	getSubDirectories : function(){
		this.checking() ;
		var ret = [] ;
		var folders = new Enumerator(
			new ActiveXObject("Scripting.FileSystemObject")
				.GetFolder( this.Path )
				.SubFolders
		) ;
		for( ; !folders.atEnd() ; folders.moveNext() ){
			ret.push(new File(folders.item().Name)) ;
		}
		
		return ret ;
	} ,	
	getSubFiles : function(){
		this.checking() ;
		var ret = [] ;
		var files = new Enumerator(
			new ActiveXObject("Scripting.FileSystemObject")
				.GetFolder(this.Path)
				.files
		)
		for( ; !files.atEnd() ; files.moveNext() ){
			ret.push(new Directory(files.item().Name) ) ;
		}
		return ret ;
	} ,	
	getParentDirectory : function(){
		return Directory( new ActiveXObject("Scripting.FileSystemObject").GetFolder(this.Path).ParentFolder )
	} ,
	copyTo : function( destination , overwrite){
		if(!overwrite) overwite = false ;
		new ActiveXObject("Scripting.FileSystemObject")
			.GetFolder(this.Path)
			.Copy(destination , overwrite) ;
	} ,
	Delete : function(){
		new ActiveXObject("Scripting.FileSystemObject")
				.GetFolder(this.Path)
				.Delete() ;
	} ,
	Move : function(destination){
		new ActiveXObject("Scripting.FileSystemObject")
				.GetFolder(this.Path)
				.Move(destination) ;
	} ,
	CreateFile : function(fileName , overwrite , encoding){
		if(!overwrite) overwrite = false ;
		if(encoding == "UTF-8") encoding = true
		else if(encoding == "Shift_JIS") encoding = false ;
		else encoding = false ;
		
		new ActiveXObject("Scripting.FileSystemObject")
			.GetFolder(this.Path)
			.CreateTextFile(fileName , overwrite , encoding) ;
	} ,
	CreateDirectory : function(folderName){
		new ActiveXObject("Scripting.FileSystemObject").CreateFolder(
			this.Path + "\\" + folderName
		) ;
	} ,
	getLastAccessed : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject")
			.GetFolder(this.Path)
			.DataLastAccessed
		
		if(!date) return null ;
		return new Date(date) ;
	} ,
	getCreatedDate : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject")
			.GetFolder(this.Path)
			.DateCreated
		
		if(!date) return null ;
		return new Date(date) ;
	} ,	
	getLastModified : function(){
		var date = new ActiveXObject("Scripting.FileSystemObject").
			GetFolder(this.Path)
			.DateLastModified
		
		return new Date(date) ;
	} ,
	getSize : function(){
		return new ActiveXObject("Scripting.FileSystemObject")
					.GetFolder(this.Path)
					.Size
	} ,
	this.toString = function(){ return "[object Directory]"}
} 