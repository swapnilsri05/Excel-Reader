package in.hopscotch.ExcelReaderV1.dto;

import java.util.Objects;

public class ResponseModel {

    private Integer code;
    private Object data;
    private String message;

    /**
     * @param code
     * @param data
     */
    public ResponseModel(Integer code, Object data, String message) {
        super();
        this.code = code;
        this.data = data;
        this.message = message;
    }

    public ResponseModel() {

    }

    /**
     * @return the code
     */
    public Integer getCode() {
        return code;
    }

    /**
     * @param code
     *            the code to set
     */
    public void setCode(Integer code) {
        this.code = code;
    }

    /**
     * @return the data
     */
    public Object getData() {
        return data;
    }

    /**
     * @param data
     *            the data to set
     */
    public void setData(Object data) {
        this.data = data;
    }

    /**
     * @return the message
     */
    public String getMessage() {
        return message;
    }

    /**
     * @param message
     *            the message to set
     */
    public void setMessage(String message) {
        this.message = message;
    }


    /*
     * (non-Javadoc)
     * 
     * @see java.lang.Object#hashCode()
     */
    @Override
    public int hashCode() {
        return Objects.hash(code, data, message);
    }

    /*
     * (non-Javadoc)
     * 
     * @see java.lang.Object#equals(java.lang.Object)
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (!(obj instanceof ResponseModel)) {
            return false;
        }
        ResponseModel other = (ResponseModel) obj;
        return Objects.equals(code, other.code) && Objects.equals(data, other.data) && Objects.equals(message, other.message);
    }

    @Override
    public String toString() {
        return "ResponseModel [code=" + code + ", data=" + data + ", message=" + message + "]";
    }
}
