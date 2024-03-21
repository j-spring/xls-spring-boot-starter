package org.jspring.xls.config;

import org.jspring.xls.domain.CellWrapper;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.jspring.xls.service.XlsxWritingService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableConfigurationProperties(XlsProperties.class)
@ConditionalOnClass(CellWrapper.class)
public class XlsConfiguration {

    @Autowired
    private XlsProperties properties;

    @Bean
    @ConditionalOnMissingBean
    public XlsxReadingService readingService() {
        return new XlsxReadingService(properties.getTemplatePath());
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

}
