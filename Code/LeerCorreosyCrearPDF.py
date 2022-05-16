import jinja2
import pdfkit
import win32com.client
import win32timezone


cantidad_correos = 10
ruta = open('./PalabrasClave.txt', 'r')
keywords = ruta.read().split(",")


def crea_pdf(ruta_template, info, ruta_css, i):
    nombre_template = ruta_template.split('/')[-1]
    ruta_template = ruta_template.replace(nombre_template, '')

    env = jinja2.Environment(loader=jinja2.FileSystemLoader(ruta_template))
    template = env.get_template(nombre_template)
    html = template.render(info)
    # print(html)

    options = {
        'page-size': 'Letter',
        'margin-top': '0.05in',
        'margin-right': '0.05in',
        'margin-bottom': '0.05in',
        'margin-left': '0.05in',
        'encoding': 'UTF-8',
        "enable-local-file-access": ""
    }
    try:
        pdfkit.from_string(html, './PDFs/outlook-'+str(i)+'.pdf',
                           css=ruta_css, options=options)
    except:
        try:
            ruta = open('./ruta_wkhtmltopdf.txt', 'r')
            config = pdfkit.configuration(wkhtmltopdf=ruta.read())
            ruta.close()

            pdfkit.from_string(html, './PDFs/outlook-'+str(i)+'.pdf',
                               css=ruta_css, options=options, configuration=config)
        except Exception as e:
            print(
                "Aun no se ha instalado el programa wkhtmltopdf o la ruta no está bien establecida")
            print(e)


def ObtenerCorreos():
    try:
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        folder = outlook.Folders[0].Folders[1]
        print("Accedio a outlook")

        for i in range(cantidad_correos):

            if(i != 0):
                print("Leyendo correo: " + str(i))
                item = folder.Items(i).Body
                # Aqui se obtendrán las palabras clave y se ordenarán los correos en base a ellos
                if any(keword in str(item) for keword in keywords):
                    sepa=str(folder.Items(i).ReceivedTime).split(" ")
                    crea_pdf("./Template/template.html", {
                        "remitente": folder.Items(i).Sender,
                        "asunto": folder.Items(i).Subject,
                        "fecha": sepa[0]+" "+str(sepa[1]).split(":")[0]+":"+str(sepa[1]).split(":")[1],
                        "cuerpo": folder.Items(i).HTMLBody

                    }, '', i)

                    print("Se creo el PDF "+str(i)+" correctamente")
                else:
                    print("El correo no contenia ninguna de las palabras clave")
    except Exception as e:
        print(e)


def ObtenerFechas_y_Asuntos():
    try:
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        folder = outlook.Folders[0].Folders[1]
        print("Accedio a outlook")
        file = open("./Fechas_y_Asuntos.txt", "w")
        for i in range(cantidad_correos):
            
            if(i != 0):
                print("Leyendo correo: " + str(i))
                item = folder.Items(i).Body
                # Aqui se obtendrán las palabras clave y se ordenarán los correos en base a ellos
                if any(keword in str(item) for keword in keywords):
                    
                    file.write("------------- correo numero: " +
                               str(i)+" -------------\n")
                    file.write("el correo de asunto: "+folder.Items(i).Subject+"\n")
                    sepa=str(folder.Items(i).ReceivedTime).split(" ")
                    file.write("se recibio en la fecha :" +sepa[0]+" "+str(sepa[1]).split(":")[0]+":"+str(sepa[1]).split(":")[1]+"\n")
                    file.write("------------- fin correo numero" +
                               str(i)+" -------------\n\n")

                    print("Se se escribio en el archivo el contenido del correo numero "+str(i)+" correctamente")
                else:
                    print("El correo no contenia ninguna de las palabras clave")
    except Exception as e:
        print(e)


def cambiar_cant_correos():
    cantidad_correos = int(input("Ingrese la cantidad de correos que desea leer: "))
    


def agregarPalabras():
    palabra = "a"
    while palabra != "exit":
        palabra = input(
            "ingrese la palabra a agregar (ingrese la palabra \"exit\" para dejar de ingresar palabras)")
        if(palabra != "exit"):
            keywords.append(palabra)


def menu():
    opcion = 0
    print("--------problemas: lalop0017999@gmail.com--------")
    print("Recuerde que si no se cuenta con el programa wkhtmltopdf no se ejecutará la opción número 2")
    print("despues de instalar el programa se necesitará agregarlo a las variables del entorno PATH del sistema o dar la ruta de instalacion del archivo la que suele ser C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe")
    print("la ruta del programa instalado deberá se cololacada en el archivo .txt \"ruta_wkhtmltopdf\"")
    print("Menu para obtener Correos y convertirlos a PDF")
    print("Las palabras clave con las que se basará el programa se tienen que configurar en el .txt llamado \"PalabrasClave\" o en su defecto ingresar de una por una en la opcion 1")
    print("La ejecucion del programa puede ser detenida en cualquier momento con ctrl+c de igual manera los archivos que lleve generados hasta el momento serán accesibles")
    print("---------------------------------------------------")
    while True:
        print("---------------------------------------------------")
        print(
            "1.- Ingresar palabras clave (esta opcion tomará en cuenta el archivo también) y las palabras ingresadas serán borradas una vez se termine la ejecucion del programa")
        print("2.- ObtenerCorreos que cuenten con las palabras clave y convertirlos a PDF con el formato encontrado en la carpeta \"Template\"")
        print("3.- Obetener solo fechas y Asuntos de los correos que cuenten con las palabras clave (la salida de esta opción será un .txt llamado \"Fechas_y_Asuntos\")")
        print("4.- Ingresar la cantidad de correos que se quieren leer (default 10) empieza por el mas reciente, es decir el correo 1 es el último correo recibido")
        print("---------------------------------------------------")
        print("5.-Salir")
        opcion = input("Ingresar opcion: ")
        if(opcion == "1"):
            agregarPalabras()
        elif(opcion == "2"):
            ObtenerCorreos()
        elif(opcion == "3"):
            ObtenerFechas_y_Asuntos()
        elif(opcion == "4"):
            cambiar_cant_correos()
        elif(opcion == "5"):
            print("gracias por utilizar el lector de correos de outlook")
            break
        else:
            print("opcion no valida favor de ingresar un numero del 1-6")


menu()
evitar_salir = input("presione enter tecla para salir......")
