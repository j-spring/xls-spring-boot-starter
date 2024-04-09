package org.jspring.xls.config;

import org.jspring.xls.domain.CellWrapper;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.jspring.xls.service.XlsxTableService;
import org.jspring.xls.service.XlsxWritingService;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

@AutoConfiguration
@EnableConfigurationProperties(XlsProperties.class)
//@ConditionalOnClass(CellWrapper.class)
public class XlsConfiguration {

    @Bean
    @ConditionalOnMissingBean
    public XlsxReadingService readingService(XlsProperties properties) {
        return new XlsxReadingService(properties.templatePath());
    }

    @Bean
    @ConditionalOnMissingBean
    public XlsxWritingService writingService() {
        return new XlsxWritingService();
    }

    @Bean
    @ConditionalOnMissingBean
    public XlsxSearchingService searchingService() {
        return new XlsxSearchingService();
    }

    @Bean
    @ConditionalOnMissingBean
    public XlsxTableService tableService() {
        return new XlsxTableService();
    }
}
