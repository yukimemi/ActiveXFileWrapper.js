class TestCore extends Core

  constructor: ->
    super()


test1 = new TestCore()

son =
  name: "Junia"
  age: 5
  family: null

john =
  name: "John"
  age: 18
  family:
    son

test1.dpObject john
