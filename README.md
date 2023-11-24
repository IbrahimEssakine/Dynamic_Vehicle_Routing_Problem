# Dynamic Vehicle Routing Problem Solver using VBA and OpenSolver
# Introduction:
The DVRP is a significant challenge in transportation logistics, requiring real-time adaptation to changing conditions. This project focuses on developing solutions for DVRP using the Coin-or CBC solver and an enhanced GA. It emphasizes the importance of effective solutions in optimizing transportation systems.

# Chapter I: Theoretical Aspect - Problem Description:
DVRP involves determining optimal routes for a fleet of vehicles in dynamic conditions. Various dynamic elements, such as customer appearance and dynamic travel times, make DVRP complex. The project classifies DVRP based on dynamic elements and emphasizes the importance of understanding these elements for effective solutions.

# Chapter I: Theoretical Aspect - Model Selection:
The chosen model is DVRP, considering locally and globally available information. The mathematical modeling includes equations ensuring each customer is visited once, total demand doesn't exceed vehicle capacity, and vehicles start and finish at a single depot.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/1b2e0e5e-b543-4043-9f42-a5b95eefa069)

Typical Dynamic Vehicle Routing Problem

The objective function of the DVRP assumes the existence of a depot and n customers. The depot is numbered as 0, and the customers are numbered as 1, 2, ..., n.

C_ij represents the transportation cost from point i to point j 

, where i and j ∈{1,2,…,n,n+1,…,n+k}  ( i or j could be a vehicle or customer).

q_i represents the demand of customer i.

T represents the maximum vehicle capacity.

m represents the number of vehicles.

initially, our focus is on obtaining a static optimized solution for the Vehicle Routing Problem.

Sr : Represents All the routes crossed by all Vehicles.

Sr(k) be the collection of arcs that are crossed by the vehicule k, where k = 1, 2, ..., m.

∀ k∈{1,..,m}  Sr(k)∈Sr 

Srfi,fj (k) represent the last arc crossed where fi is the start and fj is the destination of the final arc that is crossed. 

∀ k∈{1,..,m}  〖Sr〗_(fi,fj) (k)∈Sr(k) 

〖Sr〗_(fi,fj) (k)={1 if the arc fi to fj is the last crossed arc,0 otherwise}     (0)

y_ik={1 if the delivery of customer i is completed by vehicle k,0 otherwise}     (1)

y_ikj={1 if vehicle k transports from customer i to customer j,0 otherwise}     (2)


![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/53ff741b-ed1a-4c28-a04c-c11355a34a7a)

• Equation (4), (5), and (6) guarantee that each customer is visited only once by a single vehicle.

• Equation (7) ensures that the total demand of each customer does not exceed the capacity of a vehicle on each distribution line.

• Equation (8) ensures that the distribution vehicle starts and finishes at a single depot.

• Equation (9) ensures that the visited points remain visited after a new customer is added to n.

• Equation (10) guarantees that the trajectory travelled by vehicle k according to the last visited point fj is visited.

# Chapter I: Theoretical Aspect - Exact Solution Methods:
The project employs the branch and cut algorithm within the Coin-or CBC solver for exact solutions. This approach efficiently explores the search space to find optimal solutions.

# Chapter I: Theoretical Aspect - Approximate Solution Methods:
Two heuristic approaches, Nearest Neighbor Algorithm and Tabu Search, are employed. The Genetic Algorithm, a population-based metaheuristic, is implemented outside Excel using OpenSolver and CBC for optimization.

# Chapter I: Theoretical Aspect - Difference Between Single and Population Solution Metaheuristics:
Contrasts single-solution metaheuristics, focusing on refining a single solution, with population-based metaheuristics, maintaining diverse solutions through interactions. The Genetic Algorithm is an example of a population-based metaheuristic.

# Chapter II: Chosen Software Solution

# Introduction:
This chapter presents a solution developed using VBA (Visual Basic for Applications) in Excel, integrating the OpenSolver extension. The combination of VBA and OpenSolver provides users with a customizable and powerful environment to address optimization problems within the familiar Excel interface.


# 2. Tutorial
# Step 1: Initiate Problem

Download & Install OpenSolver extension in Excel. (https://opensolver.org/installing-opensolver/)

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/6d4eb7fa-7c6f-4a45-bd17-35c328c9d739)

Set solver parameters for time limit and branch-and-bound tolerance.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/08910007-2db4-430e-9677-111df1c92ca1)

Choose CBC as the solver engine within OpenSolver.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/fc17819c-58ac-46d0-b439-173282920307)

# Step 2: Find Solution
The first step is to click on the "Create a Table" button and enter the four parameters, which include: 
1.	Number of Points: Enter the total number of points or cities in your problem. This represents the number of locations that need to be visited or serviced.
2.	Number of Trucks: Specify the number of trucks available for the distribution or routing problem. This determines how many vehicles will be used to serve the points.
3.	Capacity of Each Truck: Enter the maximum capacity or load that each truck can carry. This represents the maximum quantity of goods or items that can be transported by each vehicle.
4.	Maximum Quantity of Demand: Specify the maximum quantity of demand for each point or city. This represents the maximum quantity of goods or services required at each location.
   
![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/7bb3635c-15e0-4058-9c94-b05ca00cca1e)

the "Submit" button to set up random locations and calculate Euclidean distances.

VBA code automatically generates tables and constraints (Generate matrices (distance, boolean, flow, linking) using VBA code).

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/51fd325b-700c-44e7-a3b3-53a3ca341e83)

The Distance Matrix

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/f34bfcc8-3a61-4719-915c-a0ac0af5ebc2)

The Boolean Matrix

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/2cc996fa-38b2-4639-9461-009f6c1083e9)

The Flow Matrix

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/0f54c21a-3ec7-4709-88da-ad4720619704)

The Linking Matrix

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/84f5be93-088c-4794-b739-4a253fa0312b)

Table of demandes

Plot the graph to visualize point locations.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/bedf6409-bf91-469f-a071-40c5f6e85ac8)


Click the "Solution" button to initiate OpenSolver's search for the best solution.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/a6fd804c-6f59-4dea-8f8c-42fbe629c88a)


# Step 3: Set Served Stop-Points
Update the state of delivered goods at each point.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/51c370ec-bfa1-4407-84ad-0c27bfd1f835)

Track progress and monitor completed tasks.

# Step 4: Add New Stop-Point
In dynamic scenarios, click "Add New Client" to generate a new client with a unique location and demand.
Accommodate changing demand patterns by dynamically expanding the problem space.

# Step 5: Find the New Solution
Click "Solution" to initiate a new optimization process considering the new serviced points.
Incorporate updated information and constraints for an optimized solution.

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/42abbae0-554d-44a3-83da-b6739e623a29)

# Objective Function:

![image](https://github.com/IbrahimEssakine/Dynamic_Vehicle_Routing_Problem/assets/103626975/89a2302a-7202-4b1c-a1ab-67efbbc282d5)

# Conclusion:
The GitHub repository offers a solid solution for the Dynamic Vehicle Routing Problem using Excel VBA and OpenSolver. While effective within its tool limitations, the solution doesn't explicitly consider traffic conditions or real-time distances between points. The conclusion suggests exploring AI-powered alternatives for more accurate demand forecasting, optimized routes, and improved operational efficiency in dynamic transportation scenarios. The repository encourages further advancements to enhance route planning and overall effectiveness.
Modify parameters like the number of trucks or capacity to adjust problem settings.
Experiment with different scenarios to evaluate their impact on the solution.
Note that modifying certain parameters may require creating additional points for effective accommodation.
