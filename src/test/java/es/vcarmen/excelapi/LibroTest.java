/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package es.vcarmen.excelapi;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author angelmillanmiralles
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
     * Test of getHojas method, of class Libro.
     */
    @Test
    public void testGetHojas() throws IOException {
        System.out.println("getHojas");
        Hoja hoja1 = new Hoja();
        Hoja hoja2 = new Hoja();
        List<Hoja> expResult = new ArrayList<>();
        expResult.add(hoja1);
        expResult.add(hoja2);
        Libro instance = new Libro(expResult, "nuevo.xlsx");
        List<Hoja> result = instance.getHojas();
        assertEquals(expResult, result);
    }

    /**
     * Test of setHojas method, of class Libro.
     */
    @Test
    public void testSetHojas() {
        System.out.println("setHojas");
        Hoja hoja1 = new Hoja();
        Hoja hoja2 = new Hoja();
        List<Hoja> hojas = new ArrayList<>();
        hojas.add(hoja1);
        hojas.add(hoja2);
        Libro instance = new Libro();
        instance.setHojas(hojas);
        assertEquals(instance.getHojas(), hojas); // TODO review the generated test code and remove the default call to fail.
        //
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
        //fail("The test case is a prototype.");
    }

    /**
     * Test of setNombreArchivo method, of class Libro.
     */
    @Test
    public void testSetNombreArchivo() {
        System.out.println("setNombreArchivo");
        String nombreArchivo = "laBellota.xlsx";
        Libro instance = new Libro(nombreArchivo);
        assertEquals(nombreArchivo, instance.getNombreArchivo());
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of addHoja method, of class Libro.
     *
     * @throws es.vcarmen.excelapi.ExcelApiException
     */
    @Test
    public void testAddHoja() throws ExcelApiException {
        System.out.println("addHoja");
        int filas = 10;
        int columnas = 10;
        Hoja hoja = new Hoja("prueba", filas, columnas);
        for (int i = 0; i < filas; i++) {
            for (int j = 0; j < columnas; j++) {
                hoja.setDatos((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        Libro instance = new Libro();
        instance.addHoja(hoja);
        try {
            assertEquals(instance.indexHoja(0), hoja);
        } catch (ExcelApiException e) {
            fail("No se puede acceder a la hoja" + e);
        }

    }

    /**
     * Test of removeHoja method, of class Libro.
     */
    @Test
    public void testRemoveHoja() throws Exception {
        System.out.println("removeHoja");
        int index = 0;
        Hoja hoja = new Hoja();
        Libro instance = new Libro();
        instance.addHoja(hoja);
        Hoja result = instance.removeHoja(index);
        assertEquals(hoja, result);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of indexHoja method, of class Libro. Y comprobada en otros test
     */
//    @Test
//    public void testIndexHoja() throws Exception {
//        System.out.println("indexHoja");
//        int index = 0;
//        Libro instance = null;
//        Hoja expResult = null;
//        Hoja result = instance.indexHoja(index);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }
    /**
     * Test of load method, of class Libro.
     */
    @Test
    public void testLoad_0args() throws ExcelApiException {
        System.out.println("load");
        String filename = "Test2.xlsx";
        Libro instance = new Libro();
        instance.testExtension();
        instance.load(filename); 
    // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of load method, of class Libro.
     */
//    @Test
//    public void testLoad_String() {
//        System.out.println("load");
//        String filename = "";
//        Libro instance = null;
//        instance.load(filename);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }

    /**
     * Test of save method, of class Libro.
     */
    @Test
    public void testSave_0args() throws ExcelApiException {
        System.out.println("save");
        Libro instance = new Libro("Test.xlsx");
        Hoja h1 = new Hoja("pepito6", 6, 6);
        Hoja h2 = new Hoja("pepote8", 8, 8);
        for (int i = 0; i < 6; i++) {
            for (int j = 0; j < 6; j++) {
                h1.setDatos((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        for (int i = 0; i < 8; i++) {
            for (int j = 0; j < 8; j++) {
                h2.setDatos((char) ('A' + j) + " " + (i + 1), i, j);
            }
        }
        instance.addHoja(h1);
        instance.addHoja(h2);
        try {
            instance.save();
        } catch (ExcelApiException ex) {
            fail("The test case is a prototype.");
        }

        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * Test of save method, of class Libro.
     */
//    @Test
//    public void testSave_String() throws Exception {
//        System.out.println("save");
//        String filename = "";
//        Libro instance = null;
//        instance.save(filename);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }

    /**
     * Test of testExtension method, of class Libro.
     */
    @Test
    public void testTestExtension()  {
        System.out.println("testExtension");
        Libro instance = new Libro("lamarimorena.po.xlpz");
        instance.testExtension();
        assertEquals(instance.getNombreArchivo(), "lamarimorena.po.xlsx");
    }

}
