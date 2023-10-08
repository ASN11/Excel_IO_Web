package org.example.excel_io;

import com.vaadin.flow.component.page.AppShellConfigurator;
import com.vaadin.flow.component.upload.receivers.MemoryBuffer;
import com.vaadin.flow.theme.Theme;
import com.vaadin.flow.theme.lumo.Lumo;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
@Theme(variant = Lumo.DARK)
public class ExcelIoWebApplication implements AppShellConfigurator {

    public static void main(String[] args) {
        SpringApplication.run(ExcelIoWebApplication.class, args);
    }

    @Bean
    public MemoryBuffer memoryBuffer() {
        return new MemoryBuffer();
    }
}
