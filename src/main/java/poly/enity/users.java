package poly.enity;

public class users {

	
	private String username;
	private String password;
	private String fullname;
	private boolean admin;
	private String email;
	public users(String username, String password, String fullname, boolean admin, String email) {
		super();
		this.username = username;
		this.password = password;
		this.fullname = fullname;
		this.admin = admin;
		this.email = email;
	}
	public users() {
		super();
		// TODO Auto-generated constructor stub
	}
	public String getUsername() {
		return username;
	}
	public void setUsername(String username) {
		this.username = username;
	}
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getFullname() {
		return fullname;
	}
	public void setFullname(String fullname) {
		this.fullname = fullname;
	}
	public boolean isAdmin() {
		return admin;
	}
	public void setAdmin(boolean admin) {
		this.admin = admin;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	
	
	
	
}
