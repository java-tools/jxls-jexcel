package com.jxls.plus.transform.jexcel;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.SheetSettings;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableSheet;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Leonid Vysochyn
 */
public class JexcelUtil {
    public static void setCellComment(WritableCell cell, String commentText){
        WritableCellFeatures features = new WritableCellFeatures();
        features.setComment(commentText);
        cell.setCellFeatures(features);
    }

    public static void setCellComment(WritableCell cell, String commentText, double width, double height){
        WritableCellFeatures features = new WritableCellFeatures();
        features.setComment(commentText, width, height);
        cell.setCellFeatures(features);
    }

    public static byte[] imageStreamToByteArray(InputStream is) throws IOException {
        BufferedImage image = ImageIO.read(is);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    public static void copySheetProperties(Sheet src, WritableSheet dest) {
        SheetSettings destSettings = dest.getSettings();
        SheetSettings srcSettings = src.getSettings();
        destSettings.setAutomaticFormulaCalculation( srcSettings.getAutomaticFormulaCalculation() );
        destSettings.setBottomMargin( srcSettings.getBottomMargin() );
        destSettings.setCopies( srcSettings.getCopies() );
        destSettings.setDisplayZeroValues( srcSettings.getDisplayZeroValues() );
        destSettings.setFitHeight( srcSettings.getFitHeight() );
        destSettings.setFitToPages( srcSettings.getFitToPages() );
        destSettings.setFitWidth( srcSettings.getFitWidth() );
        destSettings.setFooter( srcSettings.getFooter() );
        destSettings.setFooterMargin( srcSettings.getFooterMargin() );
        destSettings.setHeader( srcSettings.getHeader() );
        destSettings.setHeaderMargin( srcSettings.getHeaderMargin() );
        destSettings.setHorizontalFreeze( srcSettings.getHorizontalFreeze() );
        destSettings.setHorizontalPrintResolution( srcSettings.getHorizontalPrintResolution() );
        destSettings.setLeftMargin( srcSettings.getLeftMargin() );
        destSettings.setNormalMagnification( srcSettings.getNormalMagnification() );
        destSettings.setOrientation( srcSettings.getOrientation() );
        destSettings.setPageBreakPreviewMagnification( srcSettings.getPageBreakPreviewMagnification() );
        destSettings.setPageBreakPreviewMode( srcSettings.getPageBreakPreviewMode() );
        destSettings.setPageOrder( srcSettings.getPageOrder() );
        destSettings.setPageStart( srcSettings.getPageStart() );
        destSettings.setPaperSize( srcSettings.getPaperSize() );
        Range srcPrintArea = srcSettings.getPrintArea();
        if( srcPrintArea != null ){
            destSettings.setPrintArea(srcPrintArea.getTopLeft().getColumn(), srcPrintArea.getTopLeft().getRow(),
                srcPrintArea.getBottomRight().getColumn(), srcPrintArea.getBottomRight().getRow());
        }
        destSettings.setPrintGridLines( srcSettings.getPrintGridLines() );
        destSettings.setPrintHeaders( srcSettings.getPrintHeaders() );
        destSettings.setRecalculateFormulasBeforeSave( srcSettings.getRecalculateFormulasBeforeSave() );
        destSettings.setRightMargin( srcSettings.getRightMargin() );
        destSettings.setScaleFactor( srcSettings.getScaleFactor() );
        destSettings.setShowGridLines( srcSettings.getShowGridLines() );
        destSettings.setTopMargin( srcSettings.getTopMargin() );
        destSettings.setVerticalFreeze( srcSettings.getVerticalFreeze() );
        destSettings.setVerticalPrintResolution( srcSettings.getVerticalPrintResolution() );
        destSettings.setZoomFactor( srcSettings.getZoomFactor());
    }
}
