import lombok.Data;

@Data
public class Product {
    public Product(int productId, String productName, double price) {
        this.productId = productId;
        this.productName = productName;
        this.price = price;
    }

    private int productId;
    private String productName;
    private double price;

    // Getters and Setters
}