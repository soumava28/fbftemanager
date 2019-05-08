package com.defecttoexcel.main.model;

public class Customer {
	 private Long id;
	  
	  private String name;
	  
	  private String address;
	  
	  private int age;
	 
	  public Customer() {
	  }
	 
	  public Customer(Long id, String name, String address, int age) {
	    this.id = id;
	    this.name = name;
	    this.address = address;
	    this.age = age;
	  }
	 
	  public Long getId() {
	    return id;
	  }
	 
	  public void setId(Long id) {
	    this.id = id;
	  }
	 
	  public String getName() {
	    return name;
	  }
	 
	  public void setName(String name) {
	    this.name = name;
	  }
	 
	  public String getAddress() {
	    return address;
	  }
	 
	  public void setAddress(String address) {
	    this.address = address;
	  }
	 
	  public int getAge() {
	    return age;
	  }
	 
	  public void setAge(int age) {
	    this.age = age;
	  }
	 
	  @Override
	  public String toString() {
	    return "Customer [id=" + id + ", name=" + name + ", address=" + address + ", age=" + age + "]";
	  }
	 
}
