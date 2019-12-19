package top.soliloquize.poi;

import lombok.Data;

/**
 * @author wb
 * @date 2019/12/18
 */
@Data
public class User {
    private String name;
    private Integer age;
    private String height;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
        this.height = height;
    }
}
