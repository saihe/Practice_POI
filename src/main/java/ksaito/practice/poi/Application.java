package ksaito.practice.poi;

import ksaito.practice.poi.service.Executor;

public class Application {
  public static void main(String[] args) {
    Executor.builder().build().run(args);
  }
}
