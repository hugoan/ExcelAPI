package com.iesvdc.acceso.excelapi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Esta clase almacena información de libros para generar ficheros excel. Un
 * libro se compone de hojas, las hojas de filas y las filas de celdas.
 *
 * @author miguelin9
 */
public class Libro {

//Atributos
    private List<Hoja> hojas;
    private String nombreArchivo;

//Constructores
    /**
     * Constructor que crea el ArrayList de hojas y crea por defecto el libro
     * con nombre "nuevo.xlsx"
     */
    public Libro() {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }

    /**
     * Constructor que crea el ArrayList de hojas y define el nombre del libro
     * por el parametro que le entra (String).
     *
     * @param nombreArchivo String que define el nombre del Libro.
     */
    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }

//Getter y Setter
    /**
     * Método que devuelve el nombre del Libro.
     *
     * @return Devuelve el nombre del Libro.
     */
    public String getNombreArchivo() {
        return nombreArchivo;
    }

    /**
     * Método que cambia el nombre del Libro.
     *
     * @param nombreArchivo String que modifica el nombre del libro.
     */
    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }

//Métodos
    /**
     * Método que añade una hoja al Libro
     *
     * @param hoja parametro de tipo Hoja que se añade al Libro.
     * @return Devuelve un boolean indicando si se ha añadido o no.
     */
    public boolean addHoja(Hoja hoja) {
        return this.hojas.add(hoja);
    }

    /**
     * Método que borra una hoja dada su posición. Comprueba que la posición sea
     * válida, sino se lanza excepción.
     *
     * @param index Indica la posición de la Hoja a borrar.
     * @return Devuelve la hoja eliminada.
     * @throws ExcelAPIException Recoge la excepción de que la posición no
     * exista.
     */
    public Hoja removeHoja(int index) throws ExcelAPIException {
        if (index < 0 || index > this.hojas.size()) {
            throw new ExcelAPIException("Libro()::removeHoja(): Posición no válida.");
        }
        return this.hojas.remove(index);
    }

    /**
     * Método que devuelve una hoja dada su posición.
     *
     * Comprueba que la posición sea válida, sino que lanza excepción.
     *
     * @param index Indica la posición de la Hoja a devolver.
     * @return Devuelve una Hoja.
     * @throws ExcelAPIException Recoge la excepción de que la posición no
     * exista.
     */
    public Hoja indexHoja(int index) throws ExcelAPIException {
        if (index < 0 || index > this.hojas.size()) {
            throw new ExcelAPIException("Libro()::indexHoja(): Posición no válida.");
        }
        return this.hojas.get(index);
    }

    /**
     * Método que carga en memoria un archivo .xlsx El nombre del archivo le
     * entra como parametro String (debe incluir la extensión).
     *
     * @param archivo String con el nombre del archivo .xlsx debe tener el
     * nombre completo con la extensión.
     * @return Devuelve un objeto HSSFWorkbook (libro excel) de la libreria POI
     * @throws com.iesvdc.acceso.excelapi.ExcelAPIException Controla que exista
     * el archivo y si hay error en la entrada y salida del XSSFWorkbook.
     */
    public XSSFWorkbook load(String archivo) throws ExcelAPIException {
        //System.out.println("estoy dentro de load");
        FileInputStream ficheroEntrada;
        try {
            ficheroEntrada = new FileInputStream(archivo);
        } catch (FileNotFoundException ex) {
            throw new ExcelAPIException("Libro()::load(): Fichero no encontrado o no valido.");
        }
        XSSFWorkbook libro;
        try {
            libro = new XSSFWorkbook(ficheroEntrada);
        } catch (IOException ex) {
            throw new ExcelAPIException("Libro()::load(): Error al crear XSSFWorkbook.");
        }
        // recorre por cada hoja
        for (int i = 0; i < libro.getNumberOfSheets(); i++) {
            // obtenemos la hoja
            //System.out.println("estoy dentro de for de hojas"+ libro.getNumberOfSheets());
            Sheet hoja = libro.getSheetAt(i);
            // Crea las filas
            for (int j = 0; j < hoja.getLastRowNum()+1; j++) {
                //System.out.println("Estoy dentro del for de filas" + hoja.getLastRowNum());
                Row fila = hoja.createRow(j);
                if (fila == null) {
                    break;
                }
                //System.out.println(fila.getLastCellNum());
                // Crea las celdas y las rellena según su tipo                
                for (int k = 0; k < hoja.getLastRowNum()+1; k++) {// ------ hay un bug en gestLastCellNum y también en cellIterator ---- por eso uso el mismo que en las filas, esto fallaría si no le pasamos un excel cuadrado (Con las mismo numero de filas que de columans)
                    //System.out.println("estoy dentro del for de celdas");
                    Cell celda = fila.createCell(k);
                    if (celda == null) {
                        break;
                    }
                    switch (celda.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            System.out.println(celda.getAddress());                          
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            System.out.println(celda.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            System.out.println(celda.getErrorCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            System.out.println(celda.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println(celda.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println(celda.getStringCellValue());
                            break;
                        default:
                            System.out.println("estoy en el default");
                            break;
                    }
                }
            }
        }
        return libro;
    }

    /**
     * Este método guarda el Libro en un archivo .xlsx Con el nombre del libro.
     *
     * @throws ExcelAPIException Controla si hay error al guardar.
     */
    public void save() throws ExcelAPIException {
        SXSSFWorkbook wb = new SXSSFWorkbook();

        for (int i = 0; i < hojas.size(); i++) {
            Sheet sh = wb.createSheet(hojas.get(i).getNombre());
            for (int j = 0; j < hojas.get(i).getFilas(); j++) {
                Row row = sh.createRow(j);
                for (int k = 0; k < hojas.get(i).getColumnas(); k++) {
                    Cell cell = row.createCell(k);
                    cell.setCellValue(hojas.get(i).getDato(j, k));
                }
            }
        }

        //probando hacer el foreach para recorrer el ArrayList (apunte)
        /*
        for (Hoja h : this.hojas) {
            Sheet sh = wb.createSheet(h.getNombre());
        }
         */
        try {
            testExtension();
            FileOutputStream out = new FileOutputStream(this.nombreArchivo);// abre un flujo de datos
            wb.write(out);
            out.close();// cierra el flujo para que no se queden datos en el buffer
        } catch (IOException ex) {
            throw new ExcelAPIException("Error al guardar el archivo");
        } finally {
            wb.dispose();//libera el espacio de memoria del libro creado ya que el libro ocupa mucho por la libreria
        }
    }

    /**
     * Este método guarda el Libro en un archivo .xlsx Con el nombre que se le
     * pase como parametro en un String.
     *
     * @param filename String que modifica el nombre del Libro que es con el que
     * se guardara el Libro.
     * @throws ExcelAPIException Controla si hay error al guardar.
     */
    public void save(String filename) throws ExcelAPIException {
        this.nombreArchivo = filename;
        this.save();
    }

    /**
     * Check para comprobar que el nombre del archivo tiene la extensión xlsx
     */
    private void testExtension() {
        String extension = this.nombreArchivo.substring(this.nombreArchivo.toCharArray().length - 5, this.nombreArchivo.toCharArray().length);
        if (!extension.equals(".xlsx")) {
            this.nombreArchivo = this.nombreArchivo + ".xlsx";
        }
    }

}// fin de clase Libro

/* try declarativo y auto cierra el flujo. funciona a partir de 1.8:
try (FileOutputStream out = new FileOutputStream(this.nombreArchivo + ".xsl")){
            wb.write(out);          
        } catch (IOException ex) {
            throw new ExcelAPIException("Error al guardar el archivo");
        } finally {
            wb.dispose();//libera el espacio de memoria del libro creado ya que el libro ocupa mucho por la libreria
        }
 
 */

