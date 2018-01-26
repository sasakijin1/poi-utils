package com.jin.commons.poi.utils;

import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

public class DigestUtils {
    private static MessageDigest messageDigest;

    public static String digestFormulaName(String formulaStr) {
        if (messageDigest == null){
            try{
                messageDigest = MessageDigest.getInstance("md5");
            }catch (NoSuchAlgorithmException e){
                return formulaStr;
            }
        }

        messageDigest.update(formulaStr.getBytes());
        return "formulaStr_" + new BigInteger(1, messageDigest.digest()).toString(16);
    }
}
