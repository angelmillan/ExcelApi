package es.vcarmen.excelapi;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 * Esta clase gertiona cada hoja de un "libro" de excel
 * @author angelmillan@me.com
 * @version 1.4
 */
public class Hoja {

    private String[][] datos;
    private String nombre;
    private int nFilas;
    private int nColumnas;

    /**
     * Constructor por defecto, crea una hoja de 5 por 5 con nombre en blanco
     */
    public Hoja() {
        this.datos = new String[5][5];
        this.nFilas = 5;
        this.nColumnas = 5;
        this.nombre = "";
    }

    /**
     * Constructor sobreescrito con parámetros
     *
     * @param nombre String nombre de la hoja
     * @param nFilas int número de filas
     * @param nColumnas int número de columnas
     */
    public Hoja(String nombre, int nFilas, int nColumnas) {
        this.datos = new String[nFilas][nColumnas];
        this.nombre = nombre;
        this.nFilas = nFilas;
        this.nColumnas = nColumnas;
    }

    /**
     * Getter para obtener datos de una fila y una columna de una hoja
     *
     * @param fila int la fila de la que queremos obtener datos
     * @param columna nt la columna de la que queremos obtener datos
     * @return dato[fila][columna]
     */
    public String getDatos(int fila, int columna) {
        // comprobar posiciones
        return datos[fila][columna];
    }
    /**
     * Setter para colocar un dato en una fila y columna concreta
     * @param dato la cadena que queremos colocar en una posición de la hoaj
     * @param fila fila de la hoja
     * @param columna columna de la hoja
     */
    public void setDatos(String dato, int fila, int columna) {
        // comprobar posiciones
        this.datos[fila][columna] = dato;
    }
    /**
     * Getter para obtener el nombre de la hoja
     * @return el nombre de la hoja actual
     */
    public String getNombre() {
        return this.nombre;
    }
    /**
     * Setter para poner nombre a una hoja
     * @param nombre a poner a la hoja
     */
    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    /**
     * Getter para ontener el número de filas de una hoja
     * @return el número de filas
     */
    public int getNFilas() {
        return nFilas;
    }
    /**
     * Getter para obtener el número de columnas de una hoja
     * @return el número de columnas de la hoja
     */
    public int getNColumnas() {
        return nColumnas;
    }
    /**
     * Método que compara dos hojas 
     * @param hoja la hoja a comparar
     * @return true si las dos hojas son iguales
     */
    public boolean compare(Hoja hoja) {
        boolean iguales = true;

        if (this.nColumnas == hoja.getNColumnas()
                && this.nFilas == hoja.getNFilas()
                && this.nombre.equals(hoja.getNombre())) {

            for (int i = 0; i < this.nFilas; i++) {
                for (int j = 0; j < this.nColumnas; j++) {
                    if (this.datos[i][j].equals((hoja.getDatos(i, j)))) {
                        iguales = false;
                        break;
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
