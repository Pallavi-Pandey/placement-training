# Session 3 Solutions: Probability & Permutations

Worked solutions for every question in the deck, in slide order. Each one uses the shortcut taught in the session.

---

## Question 1 · Placement level (Slides 4–5)

**8 students are shortlisted for the placement committee. In how many ways can a President, a Vice-President and a Secretary be chosen? In how many ways can just 3 members be chosen?**
Roles differ, so order matters: ⁸P₃ = 8 × 7 × 6 = **336**
Members only, so order doesn't matter: ⁸C₃ = 336 ÷ 3! = **56**
(Watch out: using ⁸C₃ for the first part just because both parts pick "3 from 8".)

---

## Question 2 · Placement level (Slides 7–8)

**In how many ways can the letters of the word LEADER be arranged so that the vowels always come together?**
Vowels: E, A, E. Consonants: L, D, R.
Treat the vowels as one block → units to arrange: block, L, D, R = 4 → 4! = 24
Arrange inside the block (E, A, E, with E repeated): 3! ÷ 2! = 3
Total = 24 × 3 = **72**
(Watch out: divide for the repeated E only inside the block, not in the outer 4!.)

---

## Question 3 · CAT level (Slides 9–10)

**How many 4-digit numbers with no repeated digit are divisible by 5?**
Divisible by 5 → ends in 0 or 5.
Ends in 0: first digit 9 choices (1–9), then 8, then 7 → 9 × 8 × 7 = 504
Ends in 5: first digit can't be 0 or 5 → 8 × 8 × 7 = 448
Total = 504 + 448 = **952**
(Watch out: zero is allowed in the middle but not at the start.)

---

## Question 4 · CAT level (Slides 12–13)

**10 identical laddoos are shared among 3 children so that each child gets at least one. In how many ways can this be done?**
Identical items, each person ≥ 1 → ⁿ⁻¹Cᵣ₋₁ with n = 10, r = 3 → ⁹C₂ = (9 × 8) ÷ 2 = **36**
(Watch out: 3¹⁰ counts distinct items, not identical laddoos. If zero were allowed, the answer would be ¹²C₂ = 66.)

---

## Question 5 · CAT level (Slides 14–15)

**12 points on a plane, exactly 5 of which lie on one straight line (no other 3 collinear). How many triangles can be formed?**
All selections of 3 points: ¹²C₃ = 220
Collinear selections that form no triangle: ⁵C₃ = 10
Triangles = 220 − 10 = **210**
(Watch out: forgetting to subtract the collinear selections.)

---

## Question 6 · Placement level (Slides 17–18)

**A pool has 4 Infosys, 5 TCS and 3 Wipro interns. 3 interns are picked at random. Probability that one is picked from each company?**
Favourable: 4 × 5 × 3 = 60
Total ways to pick 3 from 12: ¹²C₃ = 220
Probability = 60 ÷ 220 = **3/11**
(Watch out: 4/12 × 5/11 × 3/10 fixes one order — multiply by 3! = 6 to get the same 3/11.)

---

## Question 7 · Placement level (Slides 19–20)

**A tells the truth 70% of the time, B tells the truth 80% of the time. Probability that they contradict each other on the same fact?**
A true, B false: 0.7 × 0.2 = 0.14
A false, B true: 0.3 × 0.8 = 0.24
Sum = 0.14 + 0.24 = **0.38**
(Watch out: counting only one of the two cases.)

---

## Question 8 · CAT level (Slides 21–22)

**Three dice are thrown together. Probability that the sum of the numbers is 10?**
Total outcomes = 6 × 6 × 6 = 216
a + b + c = 10, each 1–6 → let x = a − 1 etc.: x + y + z = 7, each 0–5
Without the upper limit: ⁹C₂ = 36. Remove cases where one variable ≥ 6: 3 × 3 = 9.
Favourable = 36 − 9 = 27 → Probability = 27 ÷ 216 = **1/8**
(Watch out: listing cases by hand takes over 3 minutes and is easy to miss one.)

---

## Question 9 · CAT level (Slides 23–24)

**A and B take turns rolling a die, A first. Whoever rolls a 6 first wins. Probability that A wins?**
A wins on roll 1, 3, 5, … : 1/6 + (5/6)²(1/6) + (5/6)⁴(1/6) + …
Geometric series, first term 1/6, ratio (5/6)² = 25/36
Sum = (1/6) ÷ (1 − 25/36) = (1/6) ÷ (11/36) = **6/11**
(Watch out: 1/2 ignores the first-mover advantage — A's chance must exceed half.)

---

## Question 10 · CAT level (Slides 25–26)

**A bag has 3 red and 4 black balls. Two balls are drawn one after another without replacement. Given the second ball is red, probability the first was also red?**
P(red, red) = 3/7 × 2/6 = 6/42
P(black, red) = 4/7 × 3/6 = 12/42
P(second is red) = 6/42 + 12/42 = 18/42
P(first red | second red) = 6 ÷ 18 = **1/3**
(Quick check: once one red is known used, 2 of the remaining 6 balls are red → 2/6 = 1/3.)
