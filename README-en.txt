setup.- First you have to isntall the attached file:
"wkhtmltox-0.12.6-1.msvc2015-win64"

then yo have to install the outlook app since from there is where this program is gonna take the mails also it will download them and change the format based on the template that you loaded that is in the "template" file, also it will turn them into PDF


Step next.-After the instalation is complete you have to make sure that the route in the file "ruta_wkhtmltopdf.txt" is correct; generally is this one "/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"



Optional step.- you can modify the file "PalabrasClave.txt" so the program can search the key words in the mail body 

	Example:
		mail,microsoft,paypal
		
	with that info in the file "PalabrasClave.txt" the program will search exactly those words to have a bigger scope you can introduce the same words but capitalized etc. depends on what you're looking for

step run.- the exe file is found in dist>LeerCorreosyCrearPDF.exe you can try to run it and everything should work

Nota: the file "Fechas_y_Asuntos.txt" will erease itself each time you run the 3 option of the program