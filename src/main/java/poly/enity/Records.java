package poly.enity;

import java.util.Date;

public class Records {
	private int id;
	private int type;
	private String reason;
	private String date;
	private String staffid;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public int getType() {
		return type;
	}

	public void setType(int type) {
		this.type = type;
	}

	public String getReason() {
		return reason;
	}

	public void setReason(String reason) {
		this.reason = reason;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public String getStaffid() {
		return staffid;
	}

	public void setStaffid(String staffid) {
		this.staffid = staffid;
	}

	public Records() {
		super();
		// TODO Auto-generated constructor stub
	}

	public Records(int id, int type, String reason, String date, String staffid) {
		super();
		this.id = id;
		this.type = type;
		this.reason = reason;
		this.date = date;
		this.staffid = staffid;
	}

}
