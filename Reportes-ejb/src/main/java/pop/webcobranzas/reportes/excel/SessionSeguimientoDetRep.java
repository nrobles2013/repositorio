/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.reportes.excel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.List;
import javax.ejb.Stateless;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pop.comun.dominio.CobCdr;
import pop.comun.dominio.CobLlamadas;
import pop.comun.dominio.CobMaeSeguimiento;
import pop.comun.dominio.CobSeguimiento;
import pop.comun.dominio.CobSeguimientoDet;
import pop.comun.dominio.MaeDeposito;
import pop.comun.dominio.MaeReporte;
import pop.webcobranzas.procesos.IRepSeguimientoDet;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbSeguimientoDetRep")
public class SessionSeguimientoDetRep implements IRepSeguimientoDet {

    @Override
    public byte[] exportarReporte(List<CobMaeSeguimiento> oCobMaeSeguimientos, MaeReporte maeReporte) throws Exception {
        //HSSFWorkbook workbook = new HSSFWorkbook();
        XSSFWorkbook workbook = new XSSFWorkbook();
        //HSSFSheet sheet = workbook.createSheet();
        XSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "Datos");
        String[] headers = new String[]{
            "CFondo",
            "DFondo",
            "CInversion",
            "TInversion",
            "DInversion",
            //-- llamada
            "CMaeSeguimiento",
            "CSeguimiento",
            "CSeguimientoDet",
            "CDisposicion",
            "TFamilia",
            "DFamilia",
            "TSituacion",
            "DSituacion",
            "Detalle",
            "Usuario",
            "Fecha",
            /*/-- cdr
            "Src",
            "DName",
            "Destino",
            "FLlamada",
            "Disposicion",
            "NTiempo",*/
            //-- compromiso
            "CCompromiso",
            "FRegistro",
            "FCompromiso",
            "MCompormiso",
            "CEstado",
            "FRespuesta",
            "DRespuesta",
            //-- deposito
            "CDeposito",
            "NDeposito",
            "Banco",
            "NOperacion",
            "FDeposito",
            "MDeposito",
            "TCambio",
            "MDepositoD"
        };

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerStyle.setFont(font);

        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        //HSSFRow headerRow = sheet.createRow(0);
        XSSFRow headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            //HSSFCell cell = headerRow.createCell(i);
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header);
        }
        int i = 0;
        for (CobMaeSeguimiento cobMaeSeguimiento : oCobMaeSeguimientos) {
            for (CobSeguimiento cobSeguimiento : cobMaeSeguimiento.getCobSeguimientoList()) {
                for (CobSeguimientoDet cobSeguimientoDet : cobSeguimiento.getCobSeguimientoDetList()) {
                    if (cobSeguimientoDet instanceof CobLlamadas) {
                        CobLlamadas ll = (CobLlamadas) cobSeguimientoDet;
                        System.out.println((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId());

                        if ((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId() == 9370) {
                            System.out.println((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId());
                        }

                        // si no tiene compromiso
                        if (ll.getCobCompromiso().getCcodCompromisoId() == null) {
                            XSSFRow dataRow = sheet.createRow(i + 1);
                            dataRow.createCell(0).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getCFondoId());
                            dataRow.createCell(1).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getDFondo());
                            dataRow.createCell(2).setCellValue((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId());
                            dataRow.createCell(3).setCellValue(cobMaeSeguimiento.getMaeInversion().getcTipoInv());
                            dataRow.createCell(4).setCellValue(cobMaeSeguimiento.getMaeInversion().getCInversion());
                            //-- llamada
                            dataRow.createCell(5).setCellValue(cobMaeSeguimiento.getCmaeSeguimientoId());
                            dataRow.createCell(6).setCellValue(cobSeguimiento.getCcobSeguimientoId());
                            dataRow.createCell(7).setCellValue((int) ll.getCcodLlamadaId());
                            dataRow.createCell(8).setCellValue(ll.getCcodDisposicionId());
                            dataRow.createCell(9).setCellValue(ll.getCtipoFamiliaId());
                            dataRow.createCell(10).setCellValue(ll.getTipoFamilia().getDdescripcion());
                            dataRow.createCell(11).setCellValue(ll.getCsituacionId());
                            dataRow.createCell(12).setCellValue(ll.getTipoAccion().getDdescripcion());
                            dataRow.createCell(13).setCellValue(ll.getDdescripcion());
                            dataRow.createCell(14).setCellValue(ll.getcUsuarioAdd());
                            dataRow.createCell(15).setCellValue(ll.getfUsuarioAdd());
                            i++;
                        }
                        // si tiene compromiso
                        if (ll.getCobCompromiso().getCcodCompromisoId() != null) {
                            // si no tiene depositos asociados
                            if (ll.getCobCompromiso().getMaeDepositos().isEmpty()) {
                                XSSFRow dataRow = sheet.createRow(i + 1);
                                dataRow.createCell(0).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getCFondoId());
                                dataRow.createCell(1).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getDFondo());
                                dataRow.createCell(2).setCellValue((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId());
                                dataRow.createCell(3).setCellValue(cobMaeSeguimiento.getMaeInversion().getcTipoInv());
                                dataRow.createCell(4).setCellValue(cobMaeSeguimiento.getMaeInversion().getCInversion());
                                //-- llamada
                                dataRow.createCell(5).setCellValue(cobMaeSeguimiento.getCmaeSeguimientoId());
                                dataRow.createCell(6).setCellValue(cobSeguimiento.getCcobSeguimientoId());
                                dataRow.createCell(7).setCellValue((int) ll.getCcodLlamadaId());
                                dataRow.createCell(8).setCellValue(ll.getCcodDisposicionId());
                                dataRow.createCell(9).setCellValue(ll.getCtipoFamiliaId());
                                dataRow.createCell(10).setCellValue(ll.getTipoFamilia().getDdescripcion());
                                dataRow.createCell(11).setCellValue(ll.getCsituacionId());
                                dataRow.createCell(12).setCellValue(ll.getTipoAccion().getDdescripcion());
                                dataRow.createCell(13).setCellValue(ll.getDdescripcion());
                                dataRow.createCell(14).setCellValue(ll.getcUsuarioAdd());
                                dataRow.createCell(15).setCellValue(ll.getfUsuarioAdd());
                                //-- compromiso
                                dataRow.createCell(16).setCellValue((int) ll.getCobCompromiso().getCcodCompromisoId());
                                dataRow.createCell(17).setCellValue(ll.getCobCompromiso().getfUsuarioAdd());

                                if (ll.getCobCompromiso().getFfecha() != null) {
                                    dataRow.createCell(18).setCellValue(ll.getCobCompromiso().getFfecha());
                                } else {
                                    dataRow.createCell(18).setCellValue("");
                                }
                                dataRow.createCell(19).setCellValue(ll.getCobCompromiso().getImonto());
                                                                
                                dataRow.createCell(20).setCellValue(ll.getCobCompromiso().getEestadoId());
                                if (ll.getCobCompromiso().getFfecObs() != null) {
                                    dataRow.createCell(21).setCellValue(ll.getCobCompromiso().getFfecObs());
                                } else {
                                    dataRow.createCell(21).setCellValue("");
                                }
                                dataRow.createCell(22).setCellValue(ll.getCobCompromiso().getDrespuesta());
                                i++;
                            }
                            // si tiene depositos asociados
                            if (!ll.getCobCompromiso().getMaeDepositos().isEmpty()) {
                                for (MaeDeposito deposito : ll.getCobCompromiso().getMaeDepositos()) {
                                    XSSFRow dataRow = sheet.createRow(i + 1);
                                    dataRow.createCell(0).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getCFondoId());
                                    dataRow.createCell(1).setCellValue(cobMaeSeguimiento.getMaeInversion().getMaeFondo().getDFondo());
                                    dataRow.createCell(2).setCellValue((int) cobMaeSeguimiento.getMaeInversion().getcMaeInversionId());
                                    dataRow.createCell(3).setCellValue(cobMaeSeguimiento.getMaeInversion().getcTipoInv());
                                    dataRow.createCell(4).setCellValue(cobMaeSeguimiento.getMaeInversion().getCInversion());
                                    //-- llamada
                                    dataRow.createCell(5).setCellValue(cobMaeSeguimiento.getCmaeSeguimientoId());
                                    dataRow.createCell(6).setCellValue(cobSeguimiento.getCcobSeguimientoId());
                                    dataRow.createCell(7).setCellValue((int) ll.getCcodLlamadaId());
                                    dataRow.createCell(8).setCellValue(ll.getCcodDisposicionId());
                                    dataRow.createCell(9).setCellValue(ll.getCtipoFamiliaId());
                                    dataRow.createCell(10).setCellValue(ll.getTipoFamilia().getDdescripcion());
                                    dataRow.createCell(11).setCellValue(ll.getCsituacionId());
                                    dataRow.createCell(12).setCellValue(ll.getTipoAccion().getDdescripcion());
                                    dataRow.createCell(13).setCellValue(ll.getDdescripcion());
                                    dataRow.createCell(14).setCellValue(ll.getcUsuarioAdd());
                                    dataRow.createCell(15).setCellValue(ll.getfUsuarioAdd());
                                    //-- compromiso
                                    dataRow.createCell(16).setCellValue((int) ll.getCobCompromiso().getCcodCompromisoId());
                                    dataRow.createCell(17).setCellValue(ll.getCobCompromiso().getfUsuarioAdd());

                                    if (ll.getCobCompromiso().getFfecha() != null) {
                                        dataRow.createCell(18).setCellValue(ll.getCobCompromiso().getFfecha());
                                    } else {
                                        dataRow.createCell(18).setCellValue("");
                                    }
                                    dataRow.createCell(19).setCellValue(ll.getCobCompromiso().getImonto());
                                    dataRow.createCell(20).setCellValue(ll.getCobCompromiso().getEestadoId());
                                    if (ll.getCobCompromiso().getFfecObs() != null) {
                                        dataRow.createCell(21).setCellValue(ll.getCobCompromiso().getFfecObs());
                                    } else {
                                        dataRow.createCell(21).setCellValue("");
                                    }
                                    dataRow.createCell(22).setCellValue(ll.getCobCompromiso().getDrespuesta());
                                    //-- deposito
                                    dataRow.createCell(23).setCellValue((int)deposito.getcMaeDepositoId());
                                    dataRow.createCell(24).setCellValue(deposito.getDBcoNoperacion());
                                    dataRow.createCell(25).setCellValue("BANCO");
                                    dataRow.createCell(26).setCellValue(deposito.getDBcoNoperacion());
                                    if (deposito.getFBcoDeposito() != null) {
                                        dataRow.createCell(27).setCellValue(deposito.getFBcoDeposito());
                                    } else {
                                        dataRow.createCell(27).setCellValue("");
                                    }

                                    dataRow.createCell(28).setCellValue((float) deposito.getIBcoDepositado());
                                    dataRow.createCell(29).setCellValue(deposito.getMaeTipoCambio().getnTipoCambioVen());
                                    dataRow.createCell(30).setCellValue((float) deposito.getiBcoDepositadoD());
                                    i++;
                                }
                            }
                        }

                    }
                }
            }
        }

        //HSSFRow dataRow = sheet.createRow(1 + data.length);
        //XSSFRow dataRow = sheet.createRow(1 + data.length);
        //HSSFCell total = dataRow.createCell(1);
        // XSSFCell total = dataRow.createCell(1);
        //total.setCellType(Cell.CELL_TYPE_FORMULA);
        //total.setCellStyle(style);
        //total.setCellFormula(String.format("SUM(B2:B%d)", 1 + data.length));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
        workbook.write(byteOutputStream);

        //FileOutputStream file = new FileOutputStream("workbook.xls");
        //workbook.write(file);
        //file.close();
        return byteOutputStream.toByteArray();
    }

}
