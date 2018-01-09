(defun 	C:SLCTPT()
  (setq set1 (handent "B48ED"))
  (print set1)
  (print (entget (car (entsel))))
  
  (princ)
 )