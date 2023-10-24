package org.word2pdf.reflection;

import com.lianmed.reflection.Invoker;
import com.lianmed.reflection.MethodInvoker;

import java.lang.reflect.Method;
import java.lang.reflect.ReflectPermission;
import java.util.Collection;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class Reflector {

	public static Invoker getGetInvoker(Class<?> type, String propertyName) throws Exception {
		Method[] methods = getClassMethods(type);
		Method target = null;
		for (Method method : methods) {
			if (method.getParameterTypes().length > 0) {
				continue;
			}
			String name = method.getName();
			if ((name.startsWith("get") && name.length() > 3) || (name.startsWith("is") && name.length() > 2)) {
				if(propertyName.equals(methodToProperty(name))) {
					target = method;
					break;
				}
			}
		}
		if (target == null) {
			throw new Exception("There is no getter for property named '" + propertyName + "' in '" + type + "'");
		}
		return new MethodInvoker(target);
	}

	private static String methodToProperty(String name) throws Exception {
		if (name.startsWith("is")) {
			name = name.substring(2);
		} else if (name.startsWith("get") || name.startsWith("set")) {
			name = name.substring(3);
		} else {
			throw new Exception("Error parsing property name '" + name + "'.  Didn't start with 'is', 'get' or 'set'.");
		}

		if (name.length() == 1 || (name.length() > 1 && !Character.isUpperCase(name.charAt(1)))) {
			name = name.substring(0, 1).toLowerCase(Locale.ENGLISH) + name.substring(1);
		}

		return name;
	}

	private static Method[] getClassMethods(Class<?> cls) {
		Map<String, Method> uniqueMethods = new HashMap<String, Method>();
		Class<?> currentClass = cls;
		while (currentClass != null && currentClass != Object.class) {
			addUniqueMethods(uniqueMethods, currentClass.getDeclaredMethods());

			// we also need to look for interface methods -
			// because the class may be abstract
			Class<?>[] interfaces = currentClass.getInterfaces();
			for (Class<?> anInterface : interfaces) {
				addUniqueMethods(uniqueMethods, anInterface.getMethods());
			}

			currentClass = currentClass.getSuperclass();
		}

		Collection<Method> methods = uniqueMethods.values();

		return methods.toArray(new Method[methods.size()]);
	}

	private static void addUniqueMethods(Map<String, Method> uniqueMethods, Method[] methods) {
		for (Method currentMethod : methods) {
			if (!currentMethod.isBridge()) {
				String signature = getSignature(currentMethod);
				// check to see if the method is already known
				// if it is known, then an extended class must have
				// overridden a method
				if (!uniqueMethods.containsKey(signature)) {
					if (canAccessPrivateMethods()) {
						try {
							currentMethod.setAccessible(true);
						} catch (Exception e) {
							// Ignored. This is only a final precaution, nothing we can do.
						}
					}

					uniqueMethods.put(signature, currentMethod);
				}
			}
		}
	}

	private static String getSignature(Method method) {
		StringBuilder sb = new StringBuilder();
		Class<?> returnType = method.getReturnType();
		if (returnType != null) {
			sb.append(returnType.getName()).append('#');
		}
		sb.append(method.getName());
		Class<?>[] parameters = method.getParameterTypes();
		for (int i = 0; i < parameters.length; i++) {
			if (i == 0) {
				sb.append(':');
			} else {
				sb.append(',');
			}
			sb.append(parameters[i].getName());
		}
		return sb.toString();
	}

	private static boolean canAccessPrivateMethods() {
		try {
			SecurityManager securityManager = System.getSecurityManager();
			if (null != securityManager) {
				securityManager.checkPermission(new ReflectPermission("suppressAccessChecks"));
			}
		} catch (SecurityException e) {
			return false;
		}
		return true;
	}

}
