package es.vcarmen.excelapi;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Esta clase la utilizamos para redirigir las distintas excepciones que puedan 
 * originarse a traves de nuestra propia excpcion
 * @author angelmillanmiralles
 * @version 1.4
 */
public class ExcelApiException extends Exception {
    /**
     * Constructor por defecto de ExcelApiException
     */
    public ExcelApiException() {
    }
    /**
     * Constructor sobrecargado con mensaje tipo String
     * @param msg un String con el mensaje a enviar
     */
    public ExcelApiException( String msg ) {
        super("ExcelApiExcepcion:: " + msg);
    }
    }
