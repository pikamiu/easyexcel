package com.alibaba.excel.metadata;

import java.util.List;
import java.util.Map;

/**
 * <p>write info</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2019/1/2 17:21
 */
public class WriteInfo {
    /**
     * sheet name
     */
    private final String sheetName;

    /**
     * sheet title, in most cases use chinese
     */
    private final String[] title;

    /**
     * content title, correspond contentList's key
     */
    private final String[] contentTitle;

    /**
     * content data
     */
    private final List<Map<String, Object>> contentList;

    /**
     * Private constructor ensure data consistency
     */
    private WriteInfo(Builder build) {
        this.sheetName = build.sheetName;
        this.title = build.title;
        this.contentTitle = build.contentTitle;
        this.contentList = build.contentList;
    }

    public static class Builder implements Builders<WriteInfo> {
        private String sheetName;
        private String[] title;
        private String[] contentTitle;
        private List<Map<String, Object>> contentList;

        public Builder sheetName(String val) {
            this.sheetName = val;
            return this;
        }
        public Builder title(String[] val) {
            this.title = val;
            return this;
        }
        public Builder contentTitle(String[] val) {
            this.contentTitle = val;
            return this;
        }
        public Builder contentList(List<Map<String, Object>> val) {
            this.contentList = val;
            return this;
        }

        @Override
        public WriteInfo build() {
            return new WriteInfo(this);
        }
    }

    public String getSheetName() {
        return sheetName;
    }

    public String[] getTitle() {
        return title;
    }

    public String[] getContentTitle() {
        return contentTitle;
    }

    public List<Map<String, Object>> getContentList() {
        return contentList;
    }
}
