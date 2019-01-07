package com.sailvan.excel.parameter;

import java.io.OutputStream;

import com.sailvan.excel.support.ExcelTypeEnum;
import com.sailvan.excel.ExcelWriter;

/**
 * {@link ExcelWriter}
 *
 * @author jipengfei
 */
@Deprecated
public class ExcelWriteParam {

    /**
     */
    private OutputStream outputStream;

    /**
     */
    private ExcelTypeEnum type;

    public ExcelWriteParam(OutputStream outputStream, ExcelTypeEnum type) {
        this.outputStream = outputStream;
        this.type = type;

    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public ExcelTypeEnum getType() {
        return type;
    }

    public void setType(ExcelTypeEnum type) {
        this.type = type;
    }
}
