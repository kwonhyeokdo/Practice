package com.practice.practice.elasticsearch;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.time.LocalDateTime;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import co.elastic.clients.elasticsearch._types.ElasticsearchException;
import co.elastic.clients.elasticsearch._types.query_dsl.RangeQuery;
import co.elastic.clients.elasticsearch._types.query_dsl.WildcardQuery;
import co.elastic.clients.json.JsonData;

@SpringBootTest
public class ElasticsearchServiceTest {
    @Autowired
    private ElasticsearchService elasticsearchService;

    /*  [Kibana Example]
        GET dummy/_search
    */
    @Test
    public void searchAllTest() throws ElasticsearchException, IOException{
        final String index = "dummy";

        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index(index)
        );

        assertThat(searchToListMap.size()).isEqualTo(7);
    }

    /*  [Kibana Example]
        GET dummy/_search
        {
            "query": {
                "match": {
                "text_value": "테스트"
                }
            }
        }
    */
    @Test
    public void searchMatchTest() throws ElasticsearchException, IOException{
        final String index = "dummy";
        final String field = "text_value";
        final String query = "테스트";

        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index(index)
            .query(q -> q
                .match(t -> t
                    .field(field)
                    .query(query)
                )
            )
        );

        searchToListMap.forEach(item -> {
            final String value = (String)item.get(field);
            assertThat(value).contains(query);
        });
    }

    /*  [Kibana Example]
        GET dummy/_search
        {
            "query": {
                "wildcard": {
                "text_value": "테스트*"
                }
            }
        }
    */
    @Test
    public void searchWildcardTest() throws ElasticsearchException, IOException{
        final String index = "dummy";
        final String field = "text_value";
        final String value = "테스트*";

        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index(index)
            .query(q -> q
                .wildcard(w -> w
                    .field(field)
                    .value(value)
                )
            )
        );

        searchToListMap.forEach(item -> {
            final String resultValue = (String)item.get(field);
            assertThat(resultValue).contains("테스트");
        });
    }

    /*  [Kibana Example]
        참고: https://esbook.kimjmin.net/07-settings-and-mappings/7.2-mappings/7.2.3-date
            - Elasticsearch는 기본적으로 date는 ISO8601을 사용하며, UTC 기준으로 값을 저장한다.
            - ISO8601 타입을 사용하지 않을경우 Text로 저장한다.
        GET dummy/_search
        {
        "query": {
                "range": {
                    "created_date": {
                        "gte": "2024-05-21"
                    }
                }
            }
        }
    */
    @Test
    public void searchRangeTest() throws ElasticsearchException, IOException{
        final String index = "dummy";
        final String field = "created_date";
        final String gte = "2024-05-21";

        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index(index)
            .query(q -> q
                .range(r -> r
                    .field(field)
                    .gte(JsonData.of(gte))
                )
            )
        );

        searchToListMap.forEach(item -> {
            final String dateTimeStr = (String)item.get(field);
            final LocalDateTime resultValue = LocalDateTime.from(
                Instant.from(DateTimeFormatter.ISO_DATE_TIME.parse(dateTimeStr)).atZone(ZoneId.of("UTC"))
            );
            final LocalDateTime compare = LocalDateTime.from(
                Instant.from(DateTimeFormatter.ISO_DATE_TIME.parse("2024-05-21T00:00:00.000Z")).atZone(ZoneId.of("UTC"))
            );
            assertThat(resultValue).isAfterOrEqualTo(compare);
        });
    }

    /*  [Kibana Example]
        GET dummy/_search
        {
            "query": {
                "bool": {
                    "must": [
                        {
                            "wildcard": {
                                "text_value": "테스트*"
                            } 
                        },
                        {
                            "range": {
                                    "created_date": {
                                    "gte": "2024-05-01T00:00:00",
                                    "lte": "2024-05-20T12:59:59"
                                }
                            }
                        }
                    ],
                    "must_not": [
                        {
                            "match": {
                                "number_value": 2
                            }
                        }
                    ]
                }
            }
        }
    */
    @Test
    public void searchBoolTest() throws ElasticsearchException, IOException{
        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index("dummy")
            .query(q -> q
                .bool(b -> b
                    .must(List.of(
                        WildcardQuery.of(w -> w
                            .field("text_value")
                            .value("테스트*")
                        )._toQuery(),
                        RangeQuery.of(r -> r
                            .field("created_date")
                            .gte(JsonData.of("2024-05-01T00:00:00"))
                            .lte(JsonData.of("2024-05-20T12:59:59"))
                        )._toQuery()
                    ))
                    .mustNot(mn -> mn
                        .match(m -> m
                            .field("number_value")
                            .query(2)
                        )
                    )
                )
            )    
        );

        assertThat(searchToListMap.size()).isEqualTo(2);
        assertThat(searchToListMap.get(0).get("dummy_id")).isEqualTo(1);
        assertThat(searchToListMap.get(1).get("dummy_id")).isEqualTo(3);
    }

    /*  [Kibana Example]
        GET dummy/_search
        {
            "query": {
                    "multi_match": {
                    "query": "가나다 test3",
                    "fields": ["text_value", "created_by"]
                }
            }
        }
    */
    @Test
    public void searchMultiMatchTest() throws ElasticsearchException, IOException{
        List<Map<String, Object>> searchToListMap = elasticsearchService.searchToListMap(s -> s
            .index("dummy")
            .query(q -> q
                .multiMatch(mm -> mm
                    .query("가나다 test3")
                    .fields(List.of("text_value", "created_by"))
                )
            )    
        );

        assertThat(searchToListMap.size()).isEqualTo(2);
        assertThat(searchToListMap.get(0).get("dummy_id")).isEqualTo(3);
        assertThat(searchToListMap.get(1).get("dummy_id")).isEqualTo(7);
    }

    /*  [Kibana Example]
        GET dummy/_search
        {
            "query": {
                "match": {
                "text_value": "테스트"
                }
            }
        }
    */
    @Test
    public void searchTemplateMatchTest() throws ElasticsearchException, IOException{
        final String index = "dummy";
        final String templateId = "testJsonScript";
        final String field = "text_value";
        final String query = "테스트";

        StringBuilder templateScript = new StringBuilder();
        templateScript.append("{                                          ");
        templateScript.append("     \"query\":{                           ");
        templateScript.append("         \"match\": {                      ");
        templateScript.append("             \"{{field}}\": \"{{query}}\"  ");
        templateScript.append("         }                                 ");
        templateScript.append("     }                                     ");
        templateScript.append("}                                          ");
        elasticsearchService.registTemplate(templateId, templateScript.toString());

        List<Map<String, Object>> searchToListMap = elasticsearchService.searchTemplateToListMap(s -> s
            .index(index)
            .id(templateId)
            .params("field", JsonData.of(field))
            .params("query", JsonData.of(query))
        );

        searchToListMap.forEach(item -> {
            final String value = (String)item.get(field);
            assertThat(value).contains(query);
        });
    }
}
