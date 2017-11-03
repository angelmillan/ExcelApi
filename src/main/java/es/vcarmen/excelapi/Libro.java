package es.vcarmen.excelapi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Esta clase gestiona Libros, compuestos por Hojas desde donde se cargan o
 * escriben ficheros Excel .xlsb
 *
 * @author Angelmillan@me.com
 * @since 1978
 * @version 1.4
 */
public class Libro {

    private List<Hoja> hojas;
    private String nombreArchivo;

    /**
     * Constructor por defecto de Libro, crea una Lista(hojas) de Hoja y le
     * asigna como nombre por defecto "nuevo.xlsx"
     */
    public Libro() {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = "nuevo.xlsx";
    }

    /**
     * Constructor sobrecargado en el que pasamos como parametro el nombre del
     * archivo
     *
     * @param nombreArchivo es el nombre del archivo
     */
    public Libro(String nombreArchivo) {
        this.hojas = new ArrayList<>();
        this.nombreArchivo = nombreArchivo;
    }

    /**
     * Contructor sobrecargado en el que pasamos un List de hoja y el nombre del
     * archivo
     *
     * @param hojas Lista de hojas para construir el Libro
     * @param nombreArchivo String con el nombre del archivo
     */
    public Libro(List<Hoja> hojas, String nombreArchivo) {

        this.hojas = hojas;
        this.nombreArchivo = nombreArchivo;
    }

    /**
     * Getter de hojas, la Lista hojas de un Libro
     *
     * @return hojas una List de Hoja del Libro
     */
    public List<Hoja> getHojas() {

        return hojas;
    }

    /**
     * Setter de hojas, coloca una Lista de Hoja en un libro
     *
     * @param hojas una Lista de Hoja como hojas de un libro
     */
    public void setHojas(List<Hoja> hojas) {

        this.hojas = hojas;
    }

    /**
     * Getter para obtener el nombre de archivo de un Libro
     *
     * @return nomreArchivo el nombre del archivo
     */
    public String getNombreArchivo() {

        return nombreArchivo;
    }

    /**
     * Setter para el nombre de archivo
     *
     * @param nombreArchivo el nombre del archivo a settear
     */
    public void setNombreArchivo(String nombreArchivo) {

        this.nombreArchivo = nombreArchivo;
    }

    /**
     * Boolean que nos devuelve true si la hoja se añade correctamente a un
     * Libro
     *
     * @param hoja la hoja a añadir al libro
     * @return true si se añade correctamente la hoja, false en caso contrario
     */
    public boolean addHoja(Hoja hoja) {

        return this.hojas.add(hoja);
    }

    /**
     * Método que borra una hoja en la posición index de la lista de hojas y
     * devuelve la hoja borrada
     *
     * @param index posicion de la hoja en la List de Hoja
     * @return la hoja borrada de la posición index
     * @throws ExcelApiException donde enviamos el mensaje de error
     */
    public Hoja removeHoja(int index) throws ExcelApiException {
        // si el indice es menor que cero o mayor que el tamaño de la lista envia un mensaje de error
        if (index < 0 || index > this.hojas.size()) {
            throw new ExcelApiException("Libro::removeHoja(): Posición no válida");
        }
        return this.hojas.remove(index);
    }

    /**
     * Método para obtener la hoja que este en la posición de index en la Lista
     * de hojas del libro
     *
     * @param index posición en la lista
     * @return devuelve la hoja que está en la posición index
     * @throws ExcelApiException donde enviamos el mensaje de error
     */
    public Hoja indexHoja(int index) throws ExcelApiException {
        if (index < 0 || index > this.hojas.size()) {
            throw new ExcelApiException("Libro::indexHoja(): Posición no válida");
        }
        return this.hojas.get(index);
    }

    /**
     * Método que carga un fichero de excel en disco a las clases Hoja y Libro
     * internas
     *
     * @throws ExcelApiException donde enviamos el mensaje de error
     * @throws IOException donde enviamos el mensaje de error
     */
    public void load() throws ExcelApiException, IOException {

        //Abrimos un flujo de datos desde archivo
        FileInputStream file = new FileInputStream(new File(nombreArchivo));
        //Creamos un libro excel con el flujo de datos
        XSSFWorkbook libroExcel = new XSSFWorkbook(file);
        //Creamos una libro de la Clase interna Libro y lo inicializamos a null;
        Libro libro = new Libro();
        //Creamos una hoja de la Clase interna Hoja y lo inicializamos a null;
        Hoja hoja = null;
        // Inicializamos contadores
        int numeroDeFilas = 0;
        int numeroDeColumnas = 0;
        // inicializamos una Fila(Row) y Columna(Cell)
        Row fila = null;
        Cell celda = null;
        //Para cada una de las hojas del Libro del archivo de excel.xlsx
        for (Sheet sheet : libroExcel) {
            Iterator<Row> IteraFilas = sheet.iterator();
            // Mientras existan filas a veriguamos el numero de filas máximo de esta hoja
            while (IteraFilas.hasNext()) {
                fila = IteraFilas.next();
                numeroDeFilas++;
                Iterator<Cell> IteraCeldas = fila.cellIterator();
                numeroDeColumnas = 0;
                // mientras la fila tenga columnas averiguamos el numero de columnas máximo de esta hoja
                while (IteraCeldas.hasNext()) {
                    celda = IteraCeldas.next();
                    numeroDeColumnas++;
                }
            }
            // Creamos una Hoja con el nombre de la hoja del fichero, las filas de esta hoja y las columnas de esta hoja
            hoja = new Hoja(sheet.getSheetName(), numeroDeFilas, numeroDeColumnas);
            numeroDeFilas = 0;
            // Para cada fila de la hoja de leida del archivo
            for (Row row : sheet) {
                // para cada columna que tenga la fila del archivo
                for (Cell cell : row) {
                    // En función del tipo de dato que tenga esta celda (fila,columna)
                    // Colocamos el dato de la celda del libro leido en la misma posición de la Hoja que estamos escribiendo en el POJO
                    switch (cell.getCellTypeEnum()) {

                        case _NONE:
                            break;
                        case NUMERIC:
                            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                hoja.setDatos(cell.getDateCellValue().toString(), row.getRowNum(), cell.getColumnIndex());
                            } else {
                                hoja.setDatos(cell.getNumericCellValue() + "", row.getRowNum(), cell.getColumnIndex());
                            }
                            break;
                        case STRING:
                            hoja.setDatos(cell.getStringCellValue(), row.getRowNum(), cell.getColumnIndex());
                            break;
                        case FORMULA:
                            hoja.setDatos(cell.getCellFormula(), row.getRowNum(), cell.getColumnIndex());
                            break;
                        case BLANK:
                            hoja.setDatos("", row.getRowNum(), cell.getColumnIndex());
                            break;
                        case BOOLEAN:
                            hoja.setDatos(cell.getBooleanCellValue() + "", row.getRowNum(), cell.getColumnIndex());
                            break;
                        case ERROR:
                            hoja.setDatos(cell.getErrorCellValue() + "", row.getRowNum(), cell.getColumnIndex());
                            break;
                        default:
                            hoja.setDatos(cell.getDateCellValue().toString(), row.getRowNum(), cell.getColumnIndex());
                    }
                }
            }
            // añadimos la hoja al Libro del POJO
            libro.addHoja(hoja);
        }
    }

    /**
     * Método sobreescrito que añade el nombre del archivo excel.xlsx a leer
     *
     * @param filename nombre del archivo a leer
     * @throws ExcelApiException donde enviamos el mensaje de error
     */
    public void load(String filename) throws ExcelApiException {
        try {
            this.nombreArchivo = filename;
            this.load();
        } catch (IOException ex) {
            throw new ExcelApiException("Problemas al cargar el archivo" + filename);
        }
    }

    /**
     * Método que guarda Libro en formato Excel.xlsx
     *
     * @throws ExcelApiException donde enviamos el mensaje de error
     */
    public void save() throws ExcelApiException {
        SXSSFWorkbook wb = new SXSSFWorkbook();

        for (Hoja hoja : this.hojas) {
            Sheet sh = wb.createSheet(hoja.getNombre());
            for (int i = 0; i < hoja.getNFilas(); i++) {
                Row row = sh.createRow(i);
                for (int j = 0; j < hoja.getNColumnas(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(hoja.getDatos(i, j));
                }
            }
        }
        try {
            testExtension();
            try (FileOutputStream out = new FileOutputStream(this.nombreArchivo)) {
                wb.write(out);
            }
        } catch (IOException ex) {
            throw new ExcelApiException("Error al guardar el archivo");
        } finally {
            wb.dispose();
        }
    }

    /**
     * Método save sobrecargado, igual que save() pero con el nombre del archivo a escribir
     * @param filename el nombre del fichero a escribir
     * @throws ExcelApiException donde enviamos el mensaje de error
     */
    public void save(String filename) throws ExcelApiException {
        this.nombreArchivo = filename;
        testExtension();
        this.save();
    }
    /**
     * Método para comprobar que la extensión del archivo a escribir es correcto
     */
    public void testExtension() {

        int finExtensionActual = this.nombreArchivo.length();
        int inicioExtensionActual = this.nombreArchivo.lastIndexOf(".");
        String extensionActual = this.nombreArchivo.substring(inicioExtensionActual, finExtensionActual);
        String miExtension = ".xlsx";

        if (!extensionActual.equals(miExtension)) {
            this.nombreArchivo = this.nombreArchivo.substring(0, inicioExtensionActual) + miExtension;
        }

    }

}
