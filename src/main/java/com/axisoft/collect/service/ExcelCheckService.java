package com.axisoft.collect.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface ExcelCheckService {
    public List<String> validateFile(InputStream inputStream)  throws IOException;
}
