Antes de iniciar.- Primero tiene que instalar el programa adjunto:
"wkhtmltox-0.12.6-1.msvc2015-win64"

despues tienes que instalar la aplicacion de outloop porque de allí es de donde el programa tomará los correos, ademas los descargará y les cambiará el formato en base al template que se cargo en la carpeta "template" convirtiendolos en PDF

despues de iniciar.-Despues colocar la ruta de este programa en el archivo " ruta_wkhtmltopdf.txt " 
que generalmente es la siguiente "/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe "
simplemente revise que en su computadora el "wkhtmltopdf.exe" se encuentre 
en esta ruta si no, entonces cambie la ruta del archivo

Paso Opcional.- Puede modificar el archivo "PalabrasClave.txt" para que el programa busque
en el cuerpo del correo esas palabras clave se debe seguir el formato que ya lleva el archivo

	Ejemplo:
		correo,queso,skype,xochitl
		
	con ese contenido en el archivo, el programa procedera a buscar exactamente esas palabras
	se recomienda ingresarlas tambien en mayusculas para mayor covertura, el correo debe contener
	cualquier de las palabras clave para que este sea desplegado

Ejecutando.- Si todo salio bien entrar en la carpeta "dist" y allí se encuentra el ejecutable:
"LeerCorreosyCrearPDF"

Nota: el archivo "Fechas_y_Asuntos.txt" se borrara cada vez que se ejecute la opcion 3 del
programa asi que sera necesario hacer respaldos constantes de este si se requiere 