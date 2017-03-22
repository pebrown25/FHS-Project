  OPEN "LAMPE.NAM" FOR BINARY ACCESS READ AS #1
  OPEN "LAMPE.FIX" FOR BINARY ACCESS READ WRITE AS #2
  X$ = STRING$(50, 0) + SPACE$(50)
  SEEK #2, 101
  FOR X = 1 TO 285: PUT #2, , X$: NEXT X
  GET #1, 1, X$
  PUT #2, 1, X$
  SEEK #1, 4525
  SEEK #2, 4601
  FOR X = 46 TO 61
      GET #1, , X$
      PUT #2, , X$
      PRINT X$
  NEXT X
  SEEK #1, 6133
  SEEK #2, 14401
  FOR X = 144 TO 285
      GET #1, , X$
      PUT #2, , X$
      PRINT X$
  NEXT X
  CLOSE

