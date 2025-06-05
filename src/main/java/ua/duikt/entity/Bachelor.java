package ua.duikt.entity;

import jakarta.persistence.*;
import lombok.Data;

@Entity
@Data
public class Bachelor {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String fullName;
    private String masterFullName;
    private String topic;
}
