ActiveXFileWrapper.js
=====================
This is file wrapper class for ActiveX ( for example , Windows Scripting Host) .

Code sample
-----------
```js
(function(){
	var file = new File("C:\Users\someting\hello.txt") ;
	if(!file.exsist()) return false		//Exsist method
	file.getLastAccessed().getTime() 	//"getLastAccessed" method returns Date object!
	file.setBaseName("helloworld")		//it has "setBaseName" method!
	try{
		var str = file.read() ;			// one line!
	}catch(e){
		alert(e) ;
	}
	
})()
```