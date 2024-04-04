package org.jspring.xls.config;

import org.springframework.boot.context.properties.ConfigurationProperties;


@ConfigurationProperties(prefix = "spring.export.xlsx")
public class XlsProperties {
    private static final String DEFAULT_TEMPLATE_PATH = "src/main/resources/templates/xlsx/model.xls";

    private String templatePath;

    public XlsProperties(String templatePath) {
        this.templatePath = templatePath;
    }

    public String getTemplatePath() {
        return templatePath != null ? templatePath : DEFAULT_TEMPLATE_PATH;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

}
