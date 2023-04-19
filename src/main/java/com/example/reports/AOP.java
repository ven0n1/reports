package com.example.reports;

import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.springframework.stereotype.Component;

@Aspect
@Component
@Slf4j
public class AOP {
    @Around("execution(* com.example..*.*(..))")
    public Object around(ProceedingJoinPoint joinPoint) throws Throwable{
        long startTime = System.currentTimeMillis();
        Object retVal = joinPoint.proceed();
        long timeTaken = System.currentTimeMillis() - startTime;
        log.info("Time taken by {} is equal to {}",joinPoint, timeTaken);
        return retVal;
    }
}
