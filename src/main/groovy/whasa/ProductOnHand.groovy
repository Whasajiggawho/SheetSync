package whasa

class ProductOnHand {
    String PRODUCT
    String SKU
    int QUANTITY

    @Override
    public boolean equals(Object o) {
        if(o instanceof ProductOnHand)
        {
            return ((ProductOnHand) o).SKU.equalsIgnoreCase(this.SKU)// && ((ProductOnHand) o).PRODUCT.equalsIgnoreCase(this.PRODUCT)
        }
        return false;
    }

    @Override
    public String toString() {
        return "${SKU} - ${PRODUCT} - ${QUANTITY}"
    }
}
