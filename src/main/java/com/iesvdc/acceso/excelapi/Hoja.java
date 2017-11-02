package com.iesvdc.acceso.excelapi;

/**
 * Esta clase almacena información del texto de una hoja de cálculo.
 *
 * @author miguelin9
 */
public class Hoja {

// Atributos
    private  final String[][] datos;
    private String nombre;
    private final int nFilas;
    private final int nColumnas;

// Constructores
    /**
     * Crea una hoja de cálculo nueva con tamaño de 5 x 5.
     * Y nombre vacio.
     */
    public Hoja() {
        this.datos = new String[5][5];
        this.nombre = "";
        this.nFilas = 5;
        this.nColumnas = 5;
    }

    /**
     * Crea una hoja nueva de tamaño 'nFilas' por 'nColumnas'
     *
     * @param nFilas entero que define las filas.
     * @param nColumnas entero que define las columnas.
     */
    public Hoja(int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nFilas = nFilas;
        this.nColumnas = nColumnas;
        this.nombre = "";
    }

    /**
     * Crea una hoja nueva de nombre 'nombre' de tamaño 'nFilas' por 'nColumnas'
     *
     * @param nombre String que define el nombre de la Hoja.
     * @param nFilas entero que define las filas.
     * @param nColumnas entero que define las columnas.
     */
    public Hoja(String nombre, int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = nombre;
        this.nFilas = nFilas;
        this.nColumnas = nColumnas;
    }

// Getter y Setter
    /**
     * Devuelve el contenido de una celda en una fila.
     * 
     * @param fila Indica la posición de la fila.
     * @param columna Indica la posición de la columna.
     * @return Devuelve el contenido de la celda.
     * @throws ExcelAPIException Recoge la excepción de que la posición no exista.
     */
    public String getDato(int fila, int columna) throws ExcelAPIException {
        comprobarFC(fila, columna);
        return datos[fila][columna];
    }

    /**
     * Cambia o define el contenido de una celda.
     * 
     * @param dato String que define el contenido.
     * @param fila Indica la posición de la fila.
     * @param columna Indica la posición de la columna.
     * @throws ExcelAPIException Recoge la excepción de que la posición no exista.
     */
    public void setDato(String dato, int fila, int columna) throws ExcelAPIException {
        comprobarFC(fila, columna);
        this.datos[fila][columna] = dato;
    }

    /**
     * Devuelve el nombre de la Hoja.
     * 
     * @return Devuelve el nombre de la Hoja.
     */
    public String getNombre() {
        return nombre;
    }

    /**
     * Define el nombre de la Hoja.
     * 
     * @param nombre String que define el nombre de la Hoja.
     */
    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    // get para filas y columnas y el set no lo hacemos ya que no se puede redimensionar un array
    /**
     * Devuelve el número de filas de la hoja.
     * 
     * @return Devuelve el número total de filas.
     */
    public int getFilas() {
        return nFilas;
    }

    /**
     * Devuelve el numero de columnas de la hoja.
     * 
     * @return Devuelve el número total de columnas.
     */
    public int getColumnas() {
        return nColumnas;
    }

    /**
     * Comprueba que la fila y la columna indicadas existan.
     * 
     * @param fila Indica la posición de la fila.
     * @param columna Indica la posición de la columna.
     * @throws ExcelAPIException Recoge la excepción de que la posición no exista.
     */
    public void comprobarFC(int fila, int columna) throws ExcelAPIException {
        if (fila > nFilas || columna > nColumnas || fila < 0 || columna < 0) {
            throw new ExcelAPIException("Hoja()::comprobarFC(): Posición no válida.");

        }
    }

    /**
     * Compara esta hoja con un objeto de la clase Hoja que le entre como parametro.
     * 
     * @param hoja Parametro Hoja contra el que se compara.
     * @return devuelve un boolean que indica si son iguales o no.
     * @throws ExcelAPIException Recoge la excepción de que la posición no exista.
     */
    public boolean compare(Hoja hoja) throws ExcelAPIException {
        boolean iguales = true;

        if (this.nColumnas == hoja.getColumnas()
                && this.nFilas == hoja.getFilas()
                && this.nombre.equals(hoja.getNombre())) {
            for (int i = 0; i < this.nFilas; i++) {
                for (int j = 0; j < this.nColumnas; j++) {
                    if (!this.getDato(i, j).equals(hoja.getDato(i, j))) {
                        iguales = false;
                        break; // es más eficiente usar el return pero no es correcto y se usan los dos break.
                    }
                }
                if (!iguales) {
                    break;
                }
            }
        } else {
            iguales = false;
        }

        return iguales;
    }

}

