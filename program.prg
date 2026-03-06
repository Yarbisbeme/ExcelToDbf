**=======================================================**
** Llamada a la función (Ejemplo de uso)
**=======================================================**
OPEN DATABASE C:\Eikon_Vs\Eikon\Eikon_V7_QA\Exe\datasql.dbc SHARED
m_empresa = 01
CD C:\Eikon_Vs\Eikon\Eikon_V7_QA\

DO ImportarExcelDinamico WITH "C:\Users\Eikon\Downloads\Nomina\NOMINA24.XLS", "frhwd010", "frhwh010"
RETURN 

**=======================================================**
** Procedimiento Reutilizable (Con Automatización Excel)
**=======================================================**
PROCEDURE ImportarExcelDinamico
    LPARAMETERS lcRutaExcel, lcHojaExcel, lcTablaDestino
    
    LOCAL lcAliasExcel, nCamposReales, nCamposExcel, nLimite, lcCamposSQL, i, lcQuery
    LOCAL loExcel, loLibro, lcRutaCSV
    
    lcAliasExcel = JUSTSTEM(lcRutaExcel) 
    * Definimos dónde se guardará el archivo temporal seguro
    lcRutaCSV = "C:\Eikon_Vs\Eikon\Eikon_V7_QA\" + lcAliasExcel + ".csv"

    **=======================================================**
    ** 1. Conversión OLE (El Antídoto contra el Crash)
    **=======================================================**
    WAIT WINDOW "Convirtiendo Excel mediante OLE... Por favor espere." NOWAIT
    
    * Invocamos a Microsoft Excel en segundo plano
    loExcel = CreateObject("Excel.Application")
    loExcel.Visible = .F.       && Que no se vea en pantalla
    loExcel.DisplayAlerts = .F. && Que no pida confirmaciones

    * Abrimos tu archivo problemático
    loLibro = loExcel.Workbooks.Open(lcRutaExcel)
    
    * Seleccionamos la hoja específica
    loLibro.Sheets(lcHojaExcel).Select()

    * Guardamos mágicamente como CSV (Formato 6 = xlCSV)
    loLibro.SaveAs(lcRutaCSV, 6)
    
    * Cerramos todo limpiamente
    loLibro.Close(.F.)
    loExcel.Quit()
    RELEASE loExcel
    
    WAIT CLEAR

    **=======================================================**
    ** 2. Importación Segura desde CSV
    **=======================================================**
    * Creamos un cursor temporal genérico para recibir el CSV
    CREATE CURSOR cur_csv_temp (A C(250), B C(250), C C(250), D C(250), E C(250), ;
                                F C(250), G C(250), H C(250), I C(250), J C(250), ;
                                K C(250), L C(250), M C(250), N C(250), O C(250))
                                
    SELECT cur_csv_temp
    * Importamos el CSV (Esto JAMÁS crashea FoxPro)
    APPEND FROM (lcRutaCSV) TYPE CSV
    
    * Borramos la primera fila (encabezados)
    GO TOP
    DELETE
    
    * Creamos tu cursor maestro
    SELECT * FROM (lcTablaDestino) WHERE .F. INTO CURSOR cur_tabla READWRITE 

    **=======================================================**
    ** 3. El Puente SQL Automatizado
    **=======================================================**
    nCamposReales = FCOUNT("cur_tabla")
    nCamposExcel  = FCOUNT("cur_csv_temp")

    nLimite = MIN(nCamposReales, nCamposExcel)
    lcCamposSQL = ""

    FOR i = 1 TO nLimite
        lcCampoExcel = FIELD(i, "cur_csv_temp")
        lcCampoReal  = FIELD(i, "cur_tabla")
        
        lcCamposSQL = lcCamposSQL + lcCampoExcel + " AS " + lcCampoReal
        
        IF i < nLimite
            lcCamposSQL = lcCamposSQL + ", "
        ENDIF
    ENDFOR

    lcQuery = "SELECT " + lcCamposSQL + " FROM cur_csv_temp INTO CURSOR cur_puente"
    &lcQuery

    **=======================================================**
    ** 4. Inserción Perfecta
    **=======================================================**
    SELECT (lcTablaDestino)
    APPEND FROM DBF("cur_puente")

    BROWSE TITLE "¡Misterio Resuelto! Datos importados sin Crash"
    
    * Limpieza física
    ERASE (lcRutaCSV)  && Borramos el CSV temporal para no dejar basura
    
    USE IN SELECT("cur_csv_temp")  
    USE IN SELECT("cur_tabla")
    USE IN SELECT("cur_puente")

ENDPROC

USE frhwh010
BROWSE
