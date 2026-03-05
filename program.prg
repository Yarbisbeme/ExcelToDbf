**=======================================================**
** Llamada a la función (Ejemplo de uso)
**=======================================================**
OPEN DATABASE C:\Eikon_Vs\Eikon\Eikon_V7_QA\Exe\datasql.dbc SHARED
m_empresa = 01
CD C:\Eikon_Vs\Eikon\Eikon_V7_QA\

* Aquí pasas la ruta, la hoja y la tabla destino
DO ImportarExcelDinamico WITH "C:\Users\Eikon\Downloads\Nomina\NOMINA16.XLS", "frhwd010", "frhwh010"

* ¡AQUÍ ESTÁ LA SOLUCIÓN!
* Esto detiene el script principal para que no vuelva a entrar al procedure por accidente
RETURN 


**=======================================================**
** Procedimiento Reutilizable
**=======================================================**
PROCEDURE ImportarExcelDinamico
    LPARAMETERS lcRutaExcel, lcHojaExcel, lcTablaDestino
    
    LOCAL lcAliasExcel, nCamposReales, nCamposExcel, nLimite, lcCamposSQL, i, lcQuery
    
    lcAliasExcel = JUSTSTEM(lcRutaExcel) 

    **=======================================================**
    ** 1. Importación y Cursor Maestro
    **=======================================================**
    * Se agregaron paréntesis a lcHojaExcel para forzar su lectura como variable
    IMPORT FROM (lcRutaExcel) TYPE XL8 SHEET (lcHojaExcel)
    
    SELECT * FROM (lcTablaDestino) WHERE .F. INTO CURSOR cur_tabla READWRITE 

    **=======================================================**
    ** 2. El Puente SQL Automatizado
    **=======================================================**
    nCamposReales = FCOUNT("cur_tabla")
    nCamposExcel  = FCOUNT(lcAliasExcel)

    nLimite = MIN(nCamposReales, nCamposExcel)
    lcCamposSQL = ""

    FOR i = 1 TO nLimite
        lcCampoExcel = FIELD(i, lcAliasExcel)
        lcCampoReal  = FIELD(i, "cur_tabla")
        
        lcCamposSQL = lcCamposSQL + lcCampoExcel + " AS " + lcCampoReal
        
        IF i < nLimite
            lcCamposSQL = lcCamposSQL + ", "
        ENDIF
    ENDFOR

    lcQuery = "SELECT " + lcCamposSQL + " FROM " + lcAliasExcel + " WHERE RECNO() > 1 INTO CURSOR cur_puente"
    &lcQuery

    **=======================================================**
    ** 3. Inserción Perfecta
    **=======================================================**
    SELECT (lcTablaDestino)
    APPEND FROM DBF("cur_puente")

    BROWSE TITLE "Datos importados a: " + lcTablaDestino
    
    * Limpieza de memoria
    USE IN SELECT(lcAliasExcel)  
    USE IN SELECT("cur_tabla")
    USE IN SELECT("cur_puente")

ENDPROC