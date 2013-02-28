package com.jxls.plus.transform.jexcel;

import jxl.Cell;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;

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
}
