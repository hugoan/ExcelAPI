package com.iesvdc.acceso.excelapi;

/**
 * Esta clase sera la encargada de recoger las excepciones de nuestra API
 * Tiene dos constructores para llamarla con un mensaje personalizado (String) o vacio.
 * 
 * @author miguelin9
 */
public class ExcelAPIException extends Exception {

    /**
     * Creates a new instance of <code>ExcelAPIException</code> without detail
     * message.
     */
    public ExcelAPIException() {
    }

    /**
     * Constructs an instance of <code>ExcelAPIException</code> with the
     * specified detail message.
     *
     * @param msg the detail message.
     */
    public ExcelAPIException(String msg) {
        super("ExcelAPIException::" + msg);
    }
}
