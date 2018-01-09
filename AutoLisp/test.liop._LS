(defun getspots(layer/)
  (setq file (open "H:/AutoLisp/CADexport.csv" "w"))
   
  (setq set1 (ssget "_X" (list (cons 8 layer) (cons 0 "TEXT"))))
  (setq ct 0)
  (repeat (sslength set1)
    (print
		 (strcat
		   (cdr (assoc 1 (entget (ssname set1 ct))))
		   " ,"
		   (cdr (assoc 5 (entget (ssname set1 ct))))
		   " ,"
      		 )
      		file
	   )
    (prin1 (cdr (assoc 10 (entget (ssname set1 ct)))) file)
    

    (setq ct (1+ ct))
   )
   (close file)
   (princ "done")
;;;  (print entget (ssname set1 1 )))
;;;  (princ "end")
  )
