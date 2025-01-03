## Known issues 

#### 1. Startup project update

I couldn't manage to find a reliable sink for "Startup Project" changed events. 
The Startup Project choice lives somewhere inside some binary format file in the hidden `.vs` session dir.  
Hence the way changes in the choice of Startup Project should be detected is by monitoring the `.vs` dir for changes and reacting at most once every 5 seconds. That is, assume that most sequences of operations in VS are sluggish enough that not reacting for 5 seconds in the worst case is good enough in this case. 

Currently, poll for changes every 5 seconds, regardless of file system events. ([see Section 4: State tracking](#4-state-tracking))

#### 2. UI Thread Tasks

The "canonical" way of dealing with multithreading (attempting to get tasks on the main UI thread when needed) doesn't seem reliable. `RPC_E_WRONG_THREAD` errors pop up *sometimes*, when they seemingly should not. I.e. can't guarantee work ends up on the relevant thread. Behaviour consistent throughout a session: either works or it doesn't. 

Hence an uglier, more forceful approach is taken, that may or may not mess with other VS components: switch the WNDPROC of the environment's main application window with my own. `IOleCommandTarget::Exec` then reroutes commands to the new WNDPROC (executed on the main UI thread), with `PostMessage`. If the `WM_` message is not one of mine, fall back to the original WNDPROC.

#### 3. Combo box input confirmation

Must "submit" the text for the input event to register. E.g. press "Enter" or click on an item in the MRU list. Can't just input text and navigate away. 

#### 4. State tracking 

Instead of relying on `Advise...Events` (which are a mess), or monitoring file system changes, simply set a timer for an arbitrarily coarse update frequency. Assume 5 seconds is good enough for this problem. 