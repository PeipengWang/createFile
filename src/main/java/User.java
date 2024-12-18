import lombok.Data;

@Data
public class User {
    public User(int id, String name, int age) {
        this.id = id;
        this.name = name;
        this.age = age;
    }

    private int id;
    private String name;
    private int age;

    // Getters and Setters
}

