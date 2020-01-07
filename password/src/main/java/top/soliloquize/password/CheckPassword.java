package top.soliloquize.password;


import top.soliloquze.base.StringUtils;

/**
 * @author wb
 * @date 2020/1/6
 */
public class CheckPassword {

    /**
     * 数字
     */
    private static final int NUM = 1;
    /**
     * 小写字母
     */
    private static final int SMALL_LETTER = 2;
    /**
     * 大写字母
     */
    private static final int CAPITAL_LETTER = 3;
    /**
     * 其他字符
     */
    private static final int OTHER_CHAR = 4;
    /**
     * Simple password dictionary
     */
    private final static String[] DICTIONARY = {"password", "abc123", "iloveyou", "adobe123", "123123", "sunshine",
            "1314520", "a1b2c3", "123qwe", "aaa111", "qweasd", "admin", "passwd"};

    /**
     * 验证字符类型,包括数字,大写字母,小写字母和其他字符
     */
    private static int checkCharacterType(char c) {
        if (c >= 48 && c <= 57) {
            return NUM;
        }
        if (c >= 65 && c <= 90) {
            return CAPITAL_LETTER;
        }
        if (c >= 97 && c <= 122) {
            return SMALL_LETTER;
        }
        return OTHER_CHAR;
    }

    /**
     * 根据类型统计个数
     */
    private static int countLetter(String password, int type) {
        int count = 0;
        if (null != password && password.length() > 0) {
            for (char c : password.toCharArray()) {
                if (checkCharacterType(c) == type) {
                    count++;
                }
            }
        }
        return count;
    }

    /**
     * 获取密码强度等级
     */
    public static int getPasswordLength(String password) {
        if (StringUtils.equalsNull(password)) {
            throw new IllegalArgumentException("password is empty");
        }
        int len = password.length();
        int level = 0;

        if (countLetter(password, NUM) > 0) {
            level++;
        }
        if (countLetter(password, SMALL_LETTER) > 0) {
            level++;
        }
        if (len > 4 && countLetter(password, CAPITAL_LETTER) > 0) {
            level++;
        }
        if (len > 6 && countLetter(password, OTHER_CHAR) > 0) {
            level++;
        }

        if (len > 4 && countLetter(password, NUM) > 0 && countLetter(password, SMALL_LETTER) > 0
                || countLetter(password, NUM) > 0 && countLetter(password, CAPITAL_LETTER) > 0
                || countLetter(password, NUM) > 0 && countLetter(password, OTHER_CHAR) > 0
                || countLetter(password, SMALL_LETTER) > 0 && countLetter(password, CAPITAL_LETTER) > 0
                || countLetter(password, SMALL_LETTER) > 0 && countLetter(password, OTHER_CHAR) > 0
                || countLetter(password, CAPITAL_LETTER) > 0 && countLetter(password, OTHER_CHAR) > 0) {
            level++;
        }

        if (len > 6 && countLetter(password, NUM) > 0 && countLetter(password, SMALL_LETTER) > 0
                && countLetter(password, CAPITAL_LETTER) > 0 || countLetter(password, NUM) > 0
                && countLetter(password, SMALL_LETTER) > 0 && countLetter(password, OTHER_CHAR) > 0
                || countLetter(password, NUM) > 0 && countLetter(password, CAPITAL_LETTER) > 0
                && countLetter(password, OTHER_CHAR) > 0 || countLetter(password, SMALL_LETTER) > 0
                && countLetter(password, CAPITAL_LETTER) > 0 && countLetter(password, OTHER_CHAR) > 0) {
            level++;
        }

        if (len > 8 && countLetter(password, NUM) > 0 && countLetter(password, SMALL_LETTER) > 0
                && countLetter(password, CAPITAL_LETTER) > 0 && countLetter(password, OTHER_CHAR) > 0) {
            level++;
        }

        if (len > 6 && countLetter(password, NUM) >= 3 && countLetter(password, SMALL_LETTER) >= 3
                || countLetter(password, NUM) >= 3 && countLetter(password, CAPITAL_LETTER) >= 3
                || countLetter(password, NUM) >= 3 && countLetter(password, OTHER_CHAR) >= 2
                || countLetter(password, SMALL_LETTER) >= 3 && countLetter(password, CAPITAL_LETTER) >= 3
                || countLetter(password, SMALL_LETTER) >= 3 && countLetter(password, OTHER_CHAR) >= 2
                || countLetter(password, CAPITAL_LETTER) >= 3 && countLetter(password, OTHER_CHAR) >= 2) {
            level++;
        }

        if (len > 8 && countLetter(password, NUM) >= 2 && countLetter(password, SMALL_LETTER) >= 2
                && countLetter(password, CAPITAL_LETTER) >= 2 || countLetter(password, NUM) >= 2
                && countLetter(password, SMALL_LETTER) >= 2 && countLetter(password, OTHER_CHAR) >= 2
                || countLetter(password, NUM) >= 2 && countLetter(password, CAPITAL_LETTER) >= 2
                && countLetter(password, OTHER_CHAR) >= 2 || countLetter(password, SMALL_LETTER) >= 2
                && countLetter(password, CAPITAL_LETTER) >= 2 && countLetter(password, OTHER_CHAR) >= 2) {
            level++;
        }

        if (len > 10 && countLetter(password, NUM) >= 2 && countLetter(password, SMALL_LETTER) >= 2
                && countLetter(password, CAPITAL_LETTER) >= 2 && countLetter(password, OTHER_CHAR) >= 2) {
            level++;
        }

        // 其他字符长度大于等于3,强度加1
        if (countLetter(password, OTHER_CHAR) >= 3) {
            level++;
        }
        // 其他字符长度大于等于6,强度再加1
        if (countLetter(password, OTHER_CHAR) >= 6) {
            level++;
        }

        // 长度大于12,强度加1
        if (len > 12) {
            level++;
            // 长度大于等于16,强度再加1
            if (len >= 16) {
                level++;
            }
        }

        // 全大写或全小写,强度减1
        if ("abcdefghijklmnopqrstuvwxyz".indexOf(password) > 0 || "ABCDEFGHIJKLMNOPQRSTUVWXYZ".indexOf(password) > 0) {
            level--;
        }
        // 按键盘顺序,强度减1
        if ("qwertyuiop".indexOf(password) > 0 || "asdfghjkl".indexOf(password) > 0 || "zxcvbnm".indexOf(password) > 0) {
            level--;
        }
        // 按数字顺序,强度减1
        if (StringUtils.isNumeric(password) && ("01234567890".indexOf(password) > 0 || "09876543210".indexOf(password) > 0)) {
            level--;
        }

        // 全数字或全小写或全大写,强度减1
        if (countLetter(password, NUM) == len || countLetter(password, SMALL_LETTER) == len
                || countLetter(password, CAPITAL_LETTER) == len) {
            level--;
        }

        // aaabbb
        // 规律字符
        if (len % 2 == 0) {
            String part1 = password.substring(0, len / 2);
            String part2 = password.substring(len / 2);
            if (part1.equals(part2)) {
                level--;
            }
            if (StringUtils.isCharEqual(part1) && StringUtils.isCharEqual(part2)) {
                level--;
            }
        }

        // ababab
        // 规律字符
        if (len % 3 == 0) {
            String part1 = password.substring(0, len / 3);
            String part2 = password.substring(len / 3, len / 3 * 2);
            String part3 = password.substring(len / 3 * 2);
            if (part1.equals(part2) && part2.equals(part3)) {
                level--;
            }
        }

        // 19881010 or 881010
        // 如果是日期则降低等级
        if (StringUtils.isNumeric(password) && len >= 6) {
            int year = 0;
            if (len == 8 || len == 6) {
                year = Integer.parseInt(password.substring(0, len - 4));
            }
            int size = StringUtils.sizeOfInt(year);
            int month = Integer.parseInt(password.substring(size, size + 2));
            int day = Integer.parseInt(password.substring(size + 2, len));
            if (year >= 1950 && year < 2050 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                level--;
            }
        }

        // 密码在弱密码字典中,强度减1
        if (null != DICTIONARY) {
            for (String s : DICTIONARY) {
                if (password.equals(s) || s.contains(password)) {
                    level--;
                    break;
                }
            }
        }

        // 密码长度小于等于6强度减1,小于等于4再减1,小于等于3强度为0
        if (len <= 6) {
            level--;
            if (len <= 4) {
                level--;
                if (len <= 3) {
                    level = 0;
                }
            }
        }

        // 如果字符全相等则强度为0
        if (StringUtils.isCharEqual(password)) {
            level = 0;
        }

        if (level < 0) {
            level = 0;
        }

        return level;
    }

    /**
     * 获得密码的强度等级,包括容易,中等,强,比较强,非常强
     */
    public static Level getPasswordLevel(String password) {
        int level = getPasswordLength(password);
        switch (level) {
            case 0:
            case 1:
            case 2:
            case 3:
                return Level.EASY;
            case 4:
            case 5:
            case 6:
                return Level.MIDDLE;
            case 7:
            case 8:
            case 9:
                return Level.STRONG;
            case 10:
            case 11:
            case 12:
                return Level.VERY_STRONG;
            default:
                return Level.EXTREMELY_STRONG;
        }
    }


    public static void main(String[] args) {
        System.out.println(getPasswordLevel("jks^&&^@%*(^jfls;249_ada102").value);
    }
}
