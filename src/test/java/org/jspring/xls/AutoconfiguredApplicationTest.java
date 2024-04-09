package org.jspring.xls;

import org.jspring.xls.config.XlsConfiguration;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.jspring.xls.service.XlsxTableService;
import org.jspring.xls.service.XlsxWritingService;
import org.junit.jupiter.api.Test;
import org.springframework.boot.autoconfigure.AutoConfigurations;
import org.springframework.boot.test.context.runner.ApplicationContextRunner;

import static org.assertj.core.api.Assertions.assertThat;


class AutoconfiguredApplicationTest {

    private final ApplicationContextRunner runner = new ApplicationContextRunner()
            .withConfiguration(AutoConfigurations.of(XlsConfiguration.class));

    @Test
    void shouldContainXlsBeans() {
        runner.run(context -> {
            assertThat(context).hasSingleBean(XlsxReadingService.class);
            assertThat(context).hasSingleBean(XlsxWritingService.class);
            assertThat(context).hasSingleBean(XlsxSearchingService.class);
            assertThat(context).hasSingleBean(XlsxTableService.class);
        });
    }

}
