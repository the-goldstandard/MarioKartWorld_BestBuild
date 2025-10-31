Overview:

This project uses Queueing Theory to find the optimal character-vehicle combinations for 150cc races in Mario Kart World. It calculates the probabilities of possessing different numbers of coins during a race and evaluates the expected overall speed of every possible combination. The model accounts for:

- Speed increases from terrain-specific speed attributes
- The effects of the weight attribute on how the combination's speed responds to coins
- The effects of the acceleration attribute on drift and trick boost durations and the zero-to-top-speed duration

Findings:

- There are two distinct dominant strategies for the optimal character-vehicle combinations for 150cc races in Mario Kart World
- The first dominant strategy is to maximise the speed attribute on solid terrain using a solid-focused heavyweight character (i.e. Wario or Wriggler) with the Reel Racer to maximise progress while the racer is uninterrupted
- The second dominant strategy is to rely on faster recovery when necessary and higher speed increases at the middling numbers of coins in possession to enhance the racer’s overall expected speed
- Mario Kart World players may may choose either of the two dominant strategies according to their preferred approach to the game and see similar success

How to run this repository:

1. Download all the files of this repository into one folder.
2. Open and run BestBuild.py (e.g. fn-F5) with all Excel files closed.
3. See the Python Shell and open Answers.xlsx afterwards to see the results.

The Excel files must not be manually edited.
