action=input("Seleccione una opción (inserte solo el número): \n1. Cargar un pedido de CAF para generar Excel legible. \n2. Comparar dos versiones de un pedido de CAF \n3. Cargar un CBC para generar un excel temporal con la propuesta de autorrellenado. \n4. Cargar un excel temporal para autorellenar un CbC \n")
'''
action=input("Seleccione una opción (inserte solo el número): \n1. Crear una Base de Datos (SOLO ADMIN). \n2. Cargar un CBC a la Base de Datos (SOLO ADMIN). \n3. Cargar un fichero temporal a la base de datos (SOLO ADMIN). \n4. Cargar un CBC para generar un excel temporal con la propuesta de autorrellenado. \n5. Cargar un excel temporal para autorellenar un CbC \n")
if action=="1":
    pwd=getpass.getpass("INTRODUZCA LA CONTRASEÑA: ")
    if pwd=="MCG@Trenes":
        import createDatabase
    else:
        messagebox.showinfo("Acceso Denegado","La contraseña introducida es incorrecta.")
'''