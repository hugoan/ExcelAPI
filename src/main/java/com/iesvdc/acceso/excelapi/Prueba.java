/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Clase para realizar una prueba con un main del metodo load y save de la clase Libro.
 * Se carga en memoria el archivo test.xlsx que genera el save por defecto y se guarda
 * en otro archivo que al no expecificar el nombre del archivo es nuevo.xlsx
 * 
 * @author Miguel
 */
public class Prueba {
    public static void main (String[] args){
        Libro prueba = new Libro("test.xlsx");
        try {
            prueba.load("22.xlsx");
        } catch (ExcelAPIException ex) {
            Logger.getLogger(Prueba.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            prueba.save();
        } catch (ExcelAPIException ex) {
            Logger.getLogger(Prueba.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
