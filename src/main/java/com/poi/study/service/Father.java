package com.poi.study.service;

import java.util.List;

public class Father {

    private String country;

    private String province;

    private String city;

    private List<Children> children;

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public List<Children> getChildren() {
        return children;
    }

    public void setChildren(List<Children> children) {
        this.children = children;
    }

    public static class Children {

        private String name;

        private String age;

        private String name1;
        private String name2;
        private String name3;
        private String name4;
        private String name5;
        private String name6;
        private String name7;
        private String name8;
        private String name9;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getAge() {
            return age;
        }

        public void setAge(String age) {
            this.age = age;
        }

        public String getName1() {
            return name1;
        }

        public void setName1(String name1) {
            this.name1 = name1;
        }

        public String getName2() {
            return name2;
        }

        public void setName2(String name2) {
            this.name2 = name2;
        }

        public String getName3() {
            return name3;
        }

        public void setName3(String name3) {
            this.name3 = name3;
        }

        public String getName4() {
            return name4;
        }

        public void setName4(String name4) {
            this.name4 = name4;
        }

        public String getName5() {
            return name5;
        }

        public void setName5(String name5) {
            this.name5 = name5;
        }

        public String getName6() {
            return name6;
        }

        public void setName6(String name6) {
            this.name6 = name6;
        }

        public String getName7() {
            return name7;
        }

        public void setName7(String name7) {
            this.name7 = name7;
        }

        public String getName8() {
            return name8;
        }

        public void setName8(String name8) {
            this.name8 = name8;
        }

        public String getName9() {
            return name9;
        }

        public void setName9(String name9) {
            this.name9 = name9;
        }
    }
}
