package ua.duikt.repository;

import org.hibernate.Session;
import ua.duikt.entity.Bachelor;
import ua.duikt.util.HibernateUtil;

import java.util.List;

public class BachelorRepository {

    public List<Bachelor> getAll() {
        try (Session session = HibernateUtil.getSessionFactory().openSession()) {
            return session.createQuery("from Bachelor", Bachelor.class).list();
        }
    }
}
