package org.example;

import org.example.Model.TableExample;

import java.io.FileNotFoundException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static org.example.Util.GeneradorExcel.exportExcel;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws FileNotFoundException, ParseException {
        SimpleDateFormat sdfDestino = new SimpleDateFormat("ddMMyyyy");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");
        List<String> listaCabecera = new ArrayList<>();
        List<List<Object>> listaRegistros = new ArrayList<List<Object>>();
        int indice = 1;

        List<TableExample> lista = new ArrayList<>();

        TableExample archivo1 = new TableExample(3L, "archivo3.txt", "2023-05-30");
        lista.add(archivo1);

        TableExample archivo2 = new TableExample(3L, "archivo3.txt", "2023-05-30");
        lista.add(archivo2);

        TableExample archivo3 = new TableExample(3L, "archivo3.txt", "2023-05-30");
        lista.add(archivo3);

        for (TableExample TableExample : lista) {
            List<Object> listaTmp = new ArrayList<Object>();
            Date fechaGenetacionFmtDate = sdfDestino.parse(TableExample.getFechaGeneracion());
            listaTmp.add(indice++);
            listaTmp.add(TableExample.getNombreArchivo());
            listaTmp.add(simpleDateFormat.format(fechaGenetacionFmtDate));
            listaRegistros.add(listaTmp);
        }

        listaCabecera.add("Nro");
        listaCabecera.add("Nombre Archivo");
        listaCabecera.add("Fecha Generación");

        //listaCabecera.add("Nro. Transacción");
        //listaCabecera.add("Código Respuesta");
        //listaCabecera.add("Descripción Respuesta");
        //listaCabecera.add("Entidad Destino");
        //listaCabecera.add("Entidad Origen");
        //listaCabecera.add("Indicador Cliente");
        //listaCabecera.add("Trace");
        //listaCabecera.add("Fecha Proceso");
        //listaCabecera.add("Monto");
        //listaCabecera.add("Canal");
        //listaCabecera.add("Estado");

        exportExcel("Archivo Diferecias", "Lista de Archivos de Diferencias Histórico", lista.size(), listaCabecera, listaRegistros);

    }
}
