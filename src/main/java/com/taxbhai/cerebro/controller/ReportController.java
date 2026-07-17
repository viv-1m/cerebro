package com.taxbhai.cerebro.controller;

import com.taxbhai.cerebro.service.CapitalGainsReportService;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/report")
public class ReportController {

    private final CapitalGainsReportService service;

    public ReportController(CapitalGainsReportService service) {
        this.service = service;
    }

    @PostMapping(value = "/upload", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public Map<String, List<CapitalGainsReportService.CapitalGainRow>> upload(@RequestParam("files") MultipartFile[] files) throws IOException {
        if (files == null || files.length == 0) return Map.of();
        return service.generateConsolidatedReport(Arrays.asList(files));
    }

    /**
     * Convenience endpoint: loads any XLSX files from classpath:/static and processes them.
     * Useful for running against the sample files included in the project (resources/static).
     */
    @GetMapping("/static")
    public Map<String, List<CapitalGainsReportService.CapitalGainRow>> fromStatic() throws IOException {
        PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        Resource[] resources = resolver.getResources("classpath:/static/*.xlsx");
        List<MultipartFile> files = new ArrayList<>();
        for (Resource res : resources) {
            try (InputStream is = res.getInputStream()) {
                byte[] data = is.readAllBytes();
                files.add(new InMemoryMultipartFile(res.getFilename(), data));
            }
        }
        return service.generateConsolidatedReport(files);
    }

    // Minimal MultipartFile implementation backed by byte[]
    private static class InMemoryMultipartFile implements MultipartFile {
        private final String name;
        private final byte[] data;

        InMemoryMultipartFile(String name, byte[] data) {
            this.name = name;
            this.data = data == null ? new byte[0] : data;
        }

        @Override
        public String getName() { return name; }

        @Override
        public String getOriginalFilename() { return name; }

        @Override
        public String getContentType() { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }

        @Override
        public boolean isEmpty() { return data.length == 0; }

        @Override
        public long getSize() { return data.length; }

        @Override
        public byte[] getBytes() { return data; }

        @Override
        public InputStream getInputStream() { return new ByteArrayInputStream(data); }

        @Override
        public void transferTo(File dest) throws IOException, IllegalStateException {
            try (OutputStream os = new FileOutputStream(dest)) {
                os.write(data);
            }
        }
    }
}

