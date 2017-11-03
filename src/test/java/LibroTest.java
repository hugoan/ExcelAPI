
import com.iesvdc.acceso.excelapi.ExcelAPIException;
import com.iesvdc.acceso.excelapi.Hoja;
import com.iesvdc.acceso.excelapi.Libro;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author matinal
 */
public class LibroTest {

    public LibroTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    /**
     * Test of getNombreArchivo method, of class Libro.
     */
    @Test
    public void testGetNombreArchivo() {
        System.out.println("getNombreArchivo");
        Libro instance = new Libro();
        String expResult = "nuevo.xlsx";
        String result = instance.getNombreArchivo();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    }

    /**
     * Test of setNombreArchivo method, of class Libro.
     */
    @Test
    public void testSetNombreArchivo() {
        System.out.println("setNombreArchivo");
        String nombreArchivo = "unNombre.xlsx";
        Libro instance = new Libro();
        instance.setNombreArchivo(nombreArchivo);
        assertEquals(nombreArchivo, instance.getNombreArchivo());
        // TODO review the generated test code and remove the default call to fail.
        // fail("The test case is a prototype.");
    }

    /**
     * Test of addHoja method, of class Libro.
     *
     * @throws com.iesvdc.acceso.excelapi.ExcelAPIException
     */
    @Test
    public void testAddHoja() throws ExcelAPIException {
        System.out.println("addHoja");
        int filas = 20, columnas = 30;
        Hoja hoja = new Hoja("pepe", filas, columnas);

        for (int i = 0; i < filas; i++) {
            for (int j = 0; j < columnas; j++) {
                hoja.setDato((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        Libro instance = new Libro();
        instance.addHoja(hoja);
        try {
            assertEquals(instance.indexHoja(0).compare(hoja), true);
        } catch (ExcelAPIException ex) {
            fail("No pude acceder a la hoja");
        }
    }

    /**
     * Test of removeHoja method, of class Libro.
     *
     * @throws java.lang.Exception
     */
    @Test
    public void testRemoveHoja() throws Exception {
        System.out.println("removeHoja");
        int index = 0;
        Libro libro = new Libro();
        Hoja hoja = new Hoja("Hoja 1", 5, 5);
        for (int i = 0; i < 5; i++) {
            for (int j = 0; j < 5; j++) {
                hoja.setDato((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        libro.addHoja(hoja);
        Hoja hojaRemove = libro.removeHoja(index);
        assertEquals(hojaRemove.compare(hoja), true);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /* Este test no hace falta por que este método ya lo hemos usado en otros test más arriba
     *
     * Test of indexHoja method, of class Libro.
     * @throws java.lang.Exception
     */
//    @Test
//    public void testIndexHoja() throws Exception {
//        System.out.println("indexHoja");
//        int index = 0;
//        Libro instance = new Libro();
//        Hoja expResult = null;
//        Hoja result = instance.indexHoja(index);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    } 
    /**
     * Test of load method, of class Libro.
     *
     * @throws com.iesvdc.acceso.excelapi.ExcelAPIException
     */
    @Test
    public void testLoad() throws ExcelAPIException {
        System.out.println("load");
        try {
            FileInputStream ficheroEntrada = new FileInputStream("test.xlsx");
            try {
                XSSFWorkbook libro = new XSSFWorkbook(ficheroEntrada);
                Libro prueba = new Libro();
                int contador;// contara las celdas de cada fila
                int mayor;// guardara el número de celdas de la fila que más celdas tenga
                // recorre por cada hoja
                for (int i = 0; i < libro.getNumberOfSheets(); i++) {
                    // obtenemos la hoja
                    Sheet hoja = libro.getSheetAt(i);
                    contador = 0;
                    mayor = 0;
                    //recorremos las filas
                    for (Row fila : hoja) { 
                        contador = 0;// cada fila empieza por 0 celdas
                        for (Cell cell : fila) {
                            contador++;// se incrementa por celda que haya
                        }
                        if (contador > mayor) {//comprobamos con en numero de celdas de la fila anterior y se guarda el mas alto.
                            mayor = contador;
                        }
                    }
                    Hoja miHoja = new Hoja(hoja.getSheetName(), hoja.getLastRowNum() + 1, mayor);
                    // Crea las filas
                    for (int j = 0; j < hoja.getLastRowNum()+1; j++) {
                        Row fila = hoja.createRow(j);
                        if (fila == null) {
                            break;
                        }
                        // Crea las celdas y las rellena según su tipo        
                        for (int k = 0; k < mayor; k++) {
                            Cell celda = fila.createCell(k);
                            switch (celda.getCellTypeEnum()) {
                                case BLANK:
                                    miHoja.setDato(celda.getAddress().toString(), j, k);
                                    break;
                                case BOOLEAN:
                                    miHoja.setDato(String.valueOf(celda.getBooleanCellValue()), j, k);
                                    break;
                                case ERROR:
                                    miHoja.setDato(String.valueOf(celda.getErrorCellValue()), j, k);
                                    break;
                                case FORMULA:
                                    miHoja.setDato(String.valueOf(celda.getCellFormula()), j, k);
                                    break;
                                case NUMERIC:
                                    miHoja.setDato(String.valueOf(celda.getNumericCellValue()), j, k);
                                    break;
                                case STRING:
                                    miHoja.setDato(celda.getStringCellValue(), j, k);
                                    break;
                                default:
                                    miHoja.setDato("estoy en el default", j, k);
                                    break;
                            }
                        }
                    }
                prueba.addHoja(miHoja);
                }
                prueba.save();
            } catch (IOException ex) {
                throw new ExcelAPIException("Libro()::load(): Error al crear XSSFWorkbook.");
            }
        } catch (FileNotFoundException ex) {
            throw new ExcelAPIException("Libro()::load(): Fichero no encontrado o no valido.");
        }

        //Libro instance = new Libro();
        //instance.load();
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of save method, of class Libro.
     *
     * @throws java.lang.Exception
     */
    @Test
    public void testSave() throws Exception {
        System.out.println("save");
        Libro instance = new Libro("test.xlsx");
        Hoja h1 = new Hoja("Hoja1", 6, 6);
        for (int i = 0; i < 6; i++) {
            for (int j = 0; j < 6; j++) {
                h1.setDato((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        Hoja h2 = new Hoja("Hoja2", 10, 10);
        for (int i = 0; i < 10; i++) {
            for (int j = 0; j < 10; j++) {
                h2.setDato((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        instance.addHoja(h1);
        instance.addHoja(h2);
        instance.save();

        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

}
