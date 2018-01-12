package poi.model;

import java.util.*;

public class CellSelect {

    private String[] selectArray;

    private Map<String,String> selectMap = new HashMap<String, String>();

    private Map<String,String> realMap = new HashMap<String, String>();

    public CellSelect(String key, String value, List<Map<String,Object>> selectList) {
        if (selectList != null){
            selectArray = new String[selectList.size()];
            for (int i = 0; i < selectList.size(); i++) {
                Map map = selectList.get(i);
                this.selectMap.put(map.get(value).toString(),map.get(key).toString());
                this.realMap.put(map.get(key).toString(),map.get(value).toString());
            }
        }
    }

    public CellSelect(Map map) {
        if (map != null){
            this.realMap = map;
            Set<String> keys = map.keySet();
            selectArray = new String[map.size()];
            int i = 0;
            for (Object key: map.keySet()){
                this.selectMap.put(map.get(key).toString(),key.toString());
                this.selectArray[i++] = map.get(key).toString();
            }
        }
    }

    public String[] getSelectArray(){
        return selectArray;
    }

    public String getValue(String key){
        return this.selectMap.get(key);
    }

    public String getRealValue(String key){
        return this.realMap.get(key);
    }
}
