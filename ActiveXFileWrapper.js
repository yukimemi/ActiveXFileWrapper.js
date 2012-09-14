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