pawan <- read.delim("clipboard")



for (i in 1:nrow(pawan))
{ print(i)
  dir.create(paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i]))
  if (nchar(pawan$I1[i])>0)
  {
    tryCatch({
      download.file(pawan$I1[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/1.jpg"), mode = 'wb')
    }, error = function(e) {
      cat("ERROR:",conditionMessage(e), "\n")
    })
    
  }
  if (nchar(pawan$I2[i])>0)
  {
    tryCatch({
      download.file(pawan$I2[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/2.jpg"), mode = 'wb')
    }, error = function(e) {
      cat("ERROR:",conditionMessage(e), "\n")
    })
    
  }
  if (nchar(pawan$I3[i])>0)
  {
    tryCatch({
      download.file(pawan$I3[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/3.jpg"), mode = 'wb')
    }, error = function(e) {
      cat("ERROR:",conditionMessage(e), "\n")
    })}
  
  if (nchar(pawan$I4[i])>0)
  {
    tryCatch({
      download.file(pawan$I4[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/4.jpg"), mode = 'wb')
    }, error = function(e) {
      cat("ERROR:",conditionMessage(e), "\n")
    })}
  if (nchar(pawan$I5[i])>0)
  {tryCatch({
    download.file(pawan$I5[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/5.jpg"), mode = 'wb')
  }, error = function(e) {
    cat("ERROR:",conditionMessage(e), "\n")
  })}
  if (nchar(pawan$I6[i])>0)
  {tryCatch({
    download.file(pawan$I6[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/6.jpg"), mode = 'wb')
  }, error = function(e) {
    cat("ERROR:",conditionMessage(e), "\n")
  })}
  if (nchar(pawan$I7[i])>0)
  {tryCatch({
    download.file(pawan$I7[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/7.jpg"), mode = 'wb')
  }, error = function(e) {
    cat("ERROR:",conditionMessage(e), "\n")
  })}
  if (nchar(pawan$I8[i])>0)
  {tryCatch({
    download.file(pawan$I8[i], paste0('C:/Users/harsh/Desktop/HomeDecorz/',pawan$Folder[i],"/8.jpg"), mode = 'wb')
  }, error = function(e) {
    cat("ERROR:",conditionMessage(e), "\n")
  })}
  
}
