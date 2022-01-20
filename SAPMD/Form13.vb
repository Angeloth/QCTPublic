Public Class Form13
    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'form para importar/exportar masivamente el objeto completo:
        '1. 
        'Importar:
        '1. Seleccionar la compañía objetivo
        '2. Seleccionar el template code objetivo, de 1 x 1, siempre!
        '3. Seleccionar el achivo, openfiledialog
        '4. Elegir si importa y guarda ó solo importa para evaluar
        'pero OJO, aqui como le haríamos?

        'Exportar:
        '1. Seleccionar la companía objetivo
        '2. Seleccionar el template code objetivo ó todos los objetos, checkbox!
        '3. Seleccionar el archivo, savefiledialog
        '4. 

        'Features adicionales:
        '1. Importar masivo, objetos completos, con tablas
        '2. Exportar masivo, objetos completos
        '3. Exportar masivo, una compañía completa con todos sus objetos en un solo excel, o varios
        '4. Estructura de mandantes? Definir
        '5. Funcion de clonar template, al crear uno nuevo que sea copia de algun otro creado, con todo y sus reglas (casos 56a, 56b, etc)
        '6. Evaluación masiva del objeto completo (todas las hojas/tablas)
        '7. Evaluación masiva de todos los objetos de una compañía
        '8. Comentarios en headers de como construir, o doble click y que muestre el catálogo si aplica o regla de construcción, ó dependencia, etc.
        '9. Poner colores a las columnas obligatorias(rojas), opcionales(verdes), condicionales(amarillas)
        '10. Analisis de impacto de cambios en To Be Values a objetos con registros, dependencias entre objetos, condicionales, a nivel registro
        '11. Análisis de impacto de cambios en reglas en templates/hojas de campos para registros actuales de todas las compañías, a nivel registro
        '12. Notificaciones automáticas de analisis resultante de puntos anteriores a personas interesadas
        '13. Fechas límite por template. Modulo para definir fecha límite de llenado por template y compañía
        '14. Extender perfiles:
        'Admin -Admins globales. Configura usuarios y fechas limite de llenado por template, nuevos releases
        'Viewer - Solo ver información por módulo y compañías
        'Setter - Solo modifica las reglas de los templates, condiciones, dependencias, etc. Por modulo y compañías
        'Builder/customers - Se le permite solo llenar los templates y evaluar la información  y guardarla
        '15. Funcionalidad de 'Check' de terminación de un template el cual indica que ya se pueden construir los templates dependientes del objeto
        '16. Modulo para administrar nuevos 'releases', Que se marquen las compañias y modulos y por ende los templates aplicables y fechas límite para subir la info!
        '17. Lógica para liberación de templates de acuerdo a dependencias
        '18. Visualizar la estructura de llenado completa al 'arbol' maestro de construcción!
        '19. Migrar todo a la base de de Luis Adrian
        '20. Generar tokens de autenticacion para mas seguridad
        '21. Crear 'medium installer' para realizar actualizaciones automáticas/parches
        '22. Re-Explorar PutAsync para escribir datos en batch en un nodo!
        '23. 


        'Evaluación 


    End Sub
End Class