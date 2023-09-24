import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class contador{



    static double suma=0, suma_total=0;
}

public class leer {

    public static double sumar(double numero){


        contador.suma = numero/53;

        return contador.suma;

    }

    public static void main(String[] args){

        try {
            //abrimos el XSSFWorkbook
            FileInputStream f = new FileInputStream("Reporte1.xlsx");
            XSSFWorkbook libro = new XSSFWorkbook(f);

            //seleccionamos la primera hoja
            XSSFSheet hoja = libro.getSheetAt(0);

            //Cogemos todas las filas de esa hoja
            Iterator<Row> filas = hoja.iterator();
            Iterator<Cell> celdas;

            Row fila;
            Cell celda;
            //recorremos las filas
            while (filas.hasNext()) {

                //Cogemos la siguiente fila
                fila = filas.next();

                //Cogemos todas las celdas de esa fila
                celdas = fila.cellIterator();

                //REcorremos todas las celdas
                while (celdas.hasNext()) {

                    //Cogemos la siguiente fila
                    celda = celdas.next();

                    //Segun el tipo de celda, usaremos uno u otra funcion
                    switch (celda.getCellType()) {

                        case NUMERIC:

                            contador.suma_total = sumar(celda.getNumericCellValue());
                            System.out.println(contador.suma_total);

                            break;
                        case STRING:

                            System.out.println(celda.getStringCellValue());

                            break;

                    }

                }
            }

            //cerramos el libro
            f.close();
            libro.close();

        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }





    }



}
