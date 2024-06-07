// package com.practice.practice.springdataelasticsearch;

// import java.time.LocalDateTime;

// import org.springframework.data.annotation.Id;
// import org.springframework.data.elasticsearch.annotations.Document;
// import org.springframework.data.elasticsearch.annotations.Field;
// import org.springframework.data.elasticsearch.annotations.FieldType;
// import org.springframework.data.elasticsearch.annotations.Setting;

// import lombok.AllArgsConstructor;
// import lombok.Getter;
// import lombok.NoArgsConstructor;
// import lombok.Setter;

// @Document(indexName = "dummy")
// @Setting(replicas = 0)
// @AllArgsConstructor
// @NoArgsConstructor
// @Getter
// @Setter
// public class DummyDocument {
//     @Id
//     @Field(type = FieldType.Long, name = "dummy_id")
//     private Long dummyId;

//     @Field(type = FieldType.Text, name = "text_value")
//     private String textValue;

//     @Field(type = FieldType.Float, name = "number_value")
//     private Float numberValue;

//     @Field(type = FieldType.Text, name = "created_by")
//     private String createdBy;

//     @Field(type = FieldType.Date, name = "created_date")
//     private LocalDateTime createdDate;
// }
