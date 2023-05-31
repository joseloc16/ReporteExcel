package org.example.Model;

import java.util.Date;

public class TableExample {

    private Long idArchivo;
    private String nombreArchivo;
    private String fechaGeneracion;
    private String horaGeneracion;
    private Date fechaProceso;
    private Long totalTransacciones;
    private String estado;

    public TableExample() {
    }

    public TableExample(Long idArchivo, String nombreArchivo, String fechaGeneracion) {
        this.idArchivo = idArchivo;
        this.nombreArchivo = nombreArchivo;
        this.fechaGeneracion = fechaGeneracion;
    }

    public TableExample(Long idArchivo, String nombreArchivo, String fechaGeneracion, String horaGeneracion, Date fechaProceso, Long totalTransacciones, String estado) {
        this.idArchivo = idArchivo;
        this.nombreArchivo = nombreArchivo;
        this.fechaGeneracion = fechaGeneracion;
        this.horaGeneracion = horaGeneracion;
        this.fechaProceso = fechaProceso;
        this.totalTransacciones = totalTransacciones;
        this.estado = estado;
    }

    public Long getIdArchivo() {
        return idArchivo;
    }

    public void setIdArchivo(Long idArchivo) {
        this.idArchivo = idArchivo;
    }

    public String getNombreArchivo() {
        return nombreArchivo;
    }

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }

    public String getFechaGeneracion() {
        return fechaGeneracion;
    }

    public void setFechaGeneracion(String fechaGeneracion) {
        this.fechaGeneracion = fechaGeneracion;
    }

    public String getHoraGeneracion() {
        return horaGeneracion;
    }

    public void setHoraGeneracion(String horaGeneracion) {
        this.horaGeneracion = horaGeneracion;
    }

    public Date getFechaProceso() {
        return fechaProceso;
    }

    public void setFechaProceso(Date fechaProceso) {
        this.fechaProceso = fechaProceso;
    }

    public Long getTotalTransacciones() {
        return totalTransacciones;
    }

    public void setTotalTransacciones(Long totalTransacciones) {
        this.totalTransacciones = totalTransacciones;
    }

    public String getEstado() {
        return estado;
    }

    public void setEstado(String estado) {
        this.estado = estado;
    }
}
