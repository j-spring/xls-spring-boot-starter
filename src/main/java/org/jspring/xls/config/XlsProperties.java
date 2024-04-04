package org.jspring.xls.config;

import org.springframework.boot.context.properties.ConfigurationProperties;


@ConfigurationProperties(prefix = "spring.export.xlsx")
public record XlsProperties(String templatePath) {

    public static final String DEFAULT_TEMPLATE_PATH = "src/main/resources/template/Blank.xls";

    public String templatePath() {
        return templatePath != null ? templatePath : DEFAULT_TEMPLATE_PATH;
    }
}
