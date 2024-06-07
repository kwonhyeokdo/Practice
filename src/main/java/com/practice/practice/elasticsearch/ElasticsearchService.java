package com.practice.practice.elasticsearch;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.stream.Collectors;
import java.util.function.Function;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import co.elastic.clients.elasticsearch.ElasticsearchClient;
import co.elastic.clients.elasticsearch._types.ElasticsearchException;
import co.elastic.clients.elasticsearch.core.SearchRequest;
import co.elastic.clients.elasticsearch.core.SearchTemplateRequest;
import co.elastic.clients.util.ObjectBuilder;

@Service
public class ElasticsearchService {
    @Autowired
    private ElasticsearchClient esClient;

    @SuppressWarnings({ "unchecked"})
    public List<Map<String, Object>> searchToListMap(Function<SearchRequest.Builder, ObjectBuilder<SearchRequest>> fn) throws ElasticsearchException, IOException{
        return esClient.search(fn, Map.class).hits().hits()
              .stream().map(
                    hit -> new HashMap<String, Object>(hit.source())
               ).collect(Collectors.toList());
    }

    public void registTemplate(String templateId, String source) throws IOException{
        esClient.putScript(ps -> ps
            .id(templateId)
            .script(s -> s
                .lang("mustache")
                .source(source)
            )
        );
    }

    @SuppressWarnings({ "unchecked"})
    public List<Map<String, Object>> searchTemplateToListMap(Function<SearchTemplateRequest.Builder, ObjectBuilder<SearchTemplateRequest>> fn) throws ElasticsearchException, IOException{
        return esClient.searchTemplate(fn, Map.class).hits().hits()
              .stream().map(
                hit -> new HashMap<String, Object>(hit.source())
              ).collect(Collectors.toList());
    }
}
