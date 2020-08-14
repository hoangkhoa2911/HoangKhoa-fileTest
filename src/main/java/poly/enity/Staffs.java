package poly.enity;

import java.util.Date;
import java.util.Collection;

public class Staffs {
	
	private String id;
	

	private String name;
	private boolean gender;
	private String birthday;
	private String email;
	private String departid;
	private String phone;
	private String salary;
	private String notes;
	private boolean active;

	public Staffs() {
		super();
		// TODO Auto-generated constructor stub
	}

	public String getId() {
		return id;
	}



	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public boolean isGender() {
		return gender;
	}

	public void setGender(boolean gender) {
		this.gender = gender;
	}

	public String getBirthday() {
		return birthday;
	}

	public void setBirthday(String birthday) {
		this.birthday = birthday;
	}


	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getDepartid() {
		return departid;
	}

	public void setDepartid(String departid) {
		this.departid = departid;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getSalary() {
		return salary;
	}

	public void setSalary(String salary) {
		this.salary = salary;
	}

	public String getNotes() {
		return notes;
	}

	public void setNotes(String notes) {
		this.notes = notes;
	}

	public boolean isActive() {
		return active;
	}

	public void setActive(boolean active) {
		this.active = active;
	}

	public Staffs(String id, String name, boolean gender, String birthday, String email, String departid,
			String phone, String salary, String notes, boolean active) {
		super();
		this.id = id;
		this.name = name;
		this.gender = gender;
		this.birthday = birthday;
		this.email = email;
		this.departid = departid;
		this.phone = phone;
		this.salary = salary;
		this.notes = notes;
		this.active = active;
	}

}
