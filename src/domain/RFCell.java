package domain;

public class RFCell {
	String SYS;
	String Type;
	
	String eNodeBID;
	String LocalCellID;
	String ReferenceSignalPwr;
	String IfCoastal;
	String CA;
	 
	public String getSYS() {
		return SYS;
	}
	public String getType() {
		return Type;
	}
	public String geteNodeBID() {
		return eNodeBID;
	}
	public String getLocalCellID() {
		return LocalCellID;
	}
	public String getReferenceSignalPwr() {
		return ReferenceSignalPwr;
	}
	public String getIfCoastal() {
		return IfCoastal;
	}
	public String getCA() {
		return CA;
	}
	public void setSYS(String sYS) {
		SYS = sYS;
	}
	public void setType(String type) {
		Type = type;
	}
	public void seteNodeBID(String eNodeBID) {
		this.eNodeBID = eNodeBID;
	}
	public void setLocalCellID(String localCellID) {
		LocalCellID = localCellID;
	}
	public void setReferenceSignalPwr(String referenceSignalPwr) {
		ReferenceSignalPwr = referenceSignalPwr;
	}
	public void setIfCoastal(String ifCoastal) {
		IfCoastal = ifCoastal;
	}
	public void setCA(String cA) {
		CA = cA;
	}

}
