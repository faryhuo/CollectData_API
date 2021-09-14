package com.axisoft.collect.service;

import com.axisoft.collect.entites.ComputerInfo;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

public interface ComputerCollectService {
    List<ComputerInfo> getComputerInfoList(Map<String,InputStream> inputStreams);
}
