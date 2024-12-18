#### 1. Startup project update

I couldn't manage to find a reliable sink for "Startup Project" changed events. 
The Startup Project choice lives somewhere inside some binary format file in the hidden `.vs` session dir.  
Hence the way changes in the choice of Startup Project are detected is by monitoring the `.vs` dir for changes and reacting at most once every 5 seconds. That is, assume that most sequences of operations in VS are sluggish enough that not reacting for 5 seconds in the worst case is good enough in this case. 

#### 2. UI Thread Tasks

`RPC_E_WRONG­THREAD` errors pop up *sometimes*, when they seemingly should not. Behaviour consistent throughout a session: either works or it doesn't. Just restart until you end up with a happy session...

#### 3. Combo box input confirmation

Must "submit" the text for the input event to register. E.g. press "Enter" or click on an item in the MRU list. Can't just input text and navigate away. 