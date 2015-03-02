package org.jxls.transform.jexcel;

import org.jxls.common.Context;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

/**
 * Defines an interface for a cell value which knows how to write itself to a cell
 * @author Leonid Vysochyn
 *         Date: 6/18/12
 */
public interface WritableCellValue {
    WritableCell writeToCell(WritableSheet cell, int col, int row, Context context) throws WriteException;
}
