"use strict" ;
var File , Directory ;

File = function(path){
	this.Path = path ;
	this.fso = new ActiveXObject("Scripting.FileSystemObject") ;
	this.file = this.fso.GetFile(this.Path) ;
} ;

File.prototype = {
	exists : function(){
		if(!this.fso.FileExists(this.Path)){
			return false;
		}
		return true ;
	} ,
	CopyTo : function(destination , overwrite){
		if(!overwrite){
			overwrite = false ; 
		} 
		this.file.Copy(destination , overwrite) ;
	} ,

	Delete : function(force){
		if(!force){
			force = false ;
		} 
		this.file.Delete() ;
	} ,

	Move : function(destination){
		this.file.Move(destination) ;
	} ,

	getLastAccessed : function(){
		var date = this.file.DataLastAccessed ;
		
		if(!date){
			return null ;
		} 
		return new Date(date) ;
	} ,
	
	getCreatedDate : function(){
		var date = this.file.DateCreated ;
		if(!date){
			return null ;
		} 
		return new Date(date) ;
	} ,
	getLastModified : function(){
		var date = this.file.DateLastModified ;
		return new Date(date) ;
	} ,
	getSize : function(){
		return this.file.Size ;
	} ,
	getBaseName : function(){
		return this.fso.GetBaseName(this.Path) ;
	} ,
	setBaseName : function(str){
		this.file.Name = str + "." + (this.fso.GetExtensionName(this.path)) ;
	} ,
	getExtensionName : function(){
		return this.fso.GetExtensionName(this.Path) ;
	} ,
	setExtensionName : function(str){
		this.file.Name = (this.fso.GetBaseName(this.Path)) + "." + str ;
	} ,
	getName : function(){
		return this.file.Name ;
	} ,
	getFullName : function(){
		return this.file.Path + String ;
	} ,
	setName : function(str){
		this.file.Name = str ;
	} ,
	getParentDirectory : function(){
		return new Directory( this.file.ParentFolder.Path ) ;
	} ,
	read : function(encode){
		var adodb = new ActiveXObject("ADODB.Stream") , retarray = new Array() ;
		
		if(encode){
			adodb.charset = encode ;
		}else if(!encode){
			adodb.charset = "UTF-8" ;
		} 
		
		adodb.open() ;
		adodb.loadFromFile( this.Path ) ;
		while(!adodb.EOS){
			retarray[retarray.length] = adodb.ReadText(-2) + "\n" ;
		} 
		adodb.close() ;
		return retarray.join("") ;		
	} ,
	write : function(str , encode){
		if(!this.exists()){
			this.getParentDirectory().CreateFile(this.getName()) ;
		} 
		var adodb = new ActiveXObject("ADODB.Stream") ;
		if(encode){
			adodb.charset = encode ;
		}else{
			adodb.charset = "UTF-8" ;
		}
		adodb.open() ;
		adodb.Type = 2 ;
		adodb.WriteText( str ) ;
		adodb.SaveToFile(this.Path , 2) ;
		adodb.close() ;
	}
} ;

Directory = function(path){
	if(path.substring( path.length -1 , path.length) !== "\\"){
		path += "\\" ;
	} 

	this.Path = path ;
	this.fso = new ActiveXObject("Scripting.FileSystemObject") ;
	this.exists() ;
	this.dir = this.fso.GetFolder( this.Path ) ;
	this.isroot = this.dir.IsRootFolder;
} ;
Directory.prototype = {
	exists : function(){
		if(!this.fso.FolderExists(this.Path)){
			return false ;
		} 
		return true ;
	} ,
	getSubDirectories : function(){
		var ret = new Array() , folders = new Enumerator(this.dir.SubFolders) ;
		
		for( false ; !folders.atEnd() ; folders.moveNext() ){
			ret[ret.length] = new Directory( this.getFullName() + folders.item().Name) ; 
		} 
		
		return ret ;
	} ,	
	getSubFiles : function(){
		this.exists() ;
		var ret = new Array() , files = new Enumerator(this.dir.files);
		
		for( false ; files.atEnd() !== null && files.atEnd() !== undefined ; files.moveNext() ){
			ret[ret.length] = new File(this.getFullName() + files.item().Name)  ;
		} 
		
		return ret ;
	} ,	
	getParentDirectory : function(){
		return Directory( this.dir.ParentFolder ) ;
	} ,
	
	getName : function(){
		return this.dir.Name ;
	} ,
	setName : function(str){
		this.dir.Name = str ;
	} ,
	getFullName : function(){
		var str =  this.dir.Path + String ;
		if(str.substring( str.length -1 , str.length) !== "\\"){
			str += "\\" ;
		}
		return str ;
	} ,
	copyTo : function( destination , overwrite){
		if(!overwrite){
			overwite = false ;
		}
		this.dir.Copy(destination , overwrite) ;
	} ,
	Delete : function(){
		this.dir.Delete() ;
	} ,
	Move : function(destination){
		this.dir.Move(destination) ;
	} ,
	CreateFile : function(fileName , overwrite , encoding){
		if(!overwrite){
			overwrite = false ;
		}
		if(encoding === "UTF-8"){
			encoding = true ;
		}else if(encoding === "Shift_JIS"){
			encoding = false ;
		}else{
			encoding = false ;
		}
		
		this.dir.CreateTextFile(fileName , overwrite , encoding) ;
		
		return new File(this.getFullName() + fileName) ;
		
	} ,
	CreateDirectory : function(folderName){
		this.fso.CreateFolder(this.getFullName() + folderName) ;
	} ,
	getLastAccessed : function(){
		var date = this.dir.DataLastAccessed ;
		
		if(!date){
			return null ;
		}
		return new Date(date) ;
	} ,
	getCreatedDate : function(){
		var date = this.dir.DateCreated ;
		
		if(!date){
			return null ;
		}
		return new Date(date) ;
	} ,	
	getLastModified : function(){
		var date = this.dir.DateLastModified ;
		
		return new Date(date) ;
	} ,
	getSize : function(){
		return this.dir.Size ;
	} ,
	toString : function(){ 
		return "[object Directory]" ;
	}
} ;