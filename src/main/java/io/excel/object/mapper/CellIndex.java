package io.excel.object.mapper;

import java.lang.annotation.*;

/**
 * @author woniper
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface CellIndex {

    int index();

    String message() default "";
}
