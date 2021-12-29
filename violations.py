#!/usr/bin/env python3

from dataclasses import dataclass

@dataclass
class Violation:
    colname: str
    row: int
    value: str

@dataclass
class TypeViolation(Violation):
    expected: type
    actual: type

@dataclass
class RegexViolation(Violation):
    pattern: str

@dataclass
class NonEmptyViolation(Violation):
    ...