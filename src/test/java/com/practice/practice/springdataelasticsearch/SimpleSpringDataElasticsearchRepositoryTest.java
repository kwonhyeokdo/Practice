// package com.practice.practice.springdataelasticsearch;

// import org.assertj.core.api.Assertions;
// import org.junit.jupiter.api.Test;
// import org.springframework.beans.factory.annotation.Autowired;
// import org.springframework.boot.test.context.SpringBootTest;
// import org.springframework.data.elasticsearch.core.SearchHits;

// @SpringBootTest
// public class SimpleSpringDataElasticsearchRepositoryTest {
//     @Autowired
//     private SpringDataElasticsearchRepository simpleSpringDataElasticsearchRepository;

//     @Test
//     void testFindByTextValue() {
//         SearchHits<DummyDocument> searchHits = simpleSpringDataElasticsearchRepository.findByTextValue("테스트1");
//         Assertions.assertThat(searchHits.getTotalHits()).isEqualTo(1);
//     }
// }
