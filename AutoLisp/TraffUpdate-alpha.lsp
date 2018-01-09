(defun C:UPDATETRAFFICCOUNTS( )
; ------Declare Variables -----------------------------
  (setq filepath (strcat (getvar "dwgprefix") "/AutoLisp/output.csv"))
  
  (setq file (open filepath "r"))
  
    (setq c 1)
  (while  (setq line (read-line file))
    (setq delim (vl-string-search "," line))
    (setq count (substr line 1 delim))
    (setq handle (substr line  (+ delim 2)))
    (setq old (entget (handent handle)))
    (setq new (subst (cons 1 count)  ; Substitute the text field with
		   (assoc 1 old)
		 old
	     )
    )
    (setq c (1+ c))
    (print c)
    (entmod new)

   )
  (princ "success!")
  (princ)
    

; -------Modify Attribute ------------------------------
  

; --------Display Modified Entity Properties ------------
;;;  (setq ct 0)                        ; Set ct (a counter) to 0.
;;;  (textpage)                         ; Switch to the text screen.
;;;  (princ "\nentget of last entity:")
;;;  (repeat (length entl)              ; Repeat for number of members in list:
;;;    (print (nth ct entl))            ; Output a newline, then each list member.
;;;    (setq ct (1+ ct))                ; Increments the counter by one.
;;;  )                            ; Exit quietly.
)