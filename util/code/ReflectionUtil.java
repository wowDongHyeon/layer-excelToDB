package gmx.gis.util.code;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

//Vo라는 인터페이스를 구현한 type만 들어올 수 있음. 안그러면 컴파일 에러 남.
public class ReflectionUtil<T> {

    private Method[] methods;

    public ReflectionUtil(T clz) {

        // 들어온 Vo객체의 메소드를 저장
        methods = clz.getClass().getMethods();

    }

    // 들어온 스트링 값을 가지고 메소드를 호출
    public void invokeSetMethod(T clz, String methodToInvoke, String toSet) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {

        String firstString = methodToInvoke.substring(0, 1).toUpperCase();

        String modifiedString = "set" + firstString
                + methodToInvoke.substring(1, methodToInvoke.length());

        for (Method method : methods) {
            if (method.getName().equals(modifiedString)) {

                method.invoke(clz, toSet);

            }
        }
    }

    // 들어온 스트링 값을 가지고 메소드를 호출
    public String invokeGetMethod(T clz, String methodToInvoke) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {

    	String val="";

        String firstString = methodToInvoke.substring(0, 1).toUpperCase();

        String modifiedString = "get" + firstString
                + methodToInvoke.substring(1, methodToInvoke.length());

        for (Method method : methods) {
            if (method.getName().equals(modifiedString)) {

               val= (String) method.invoke(clz);

            }
        }
        return val;
    }
}

