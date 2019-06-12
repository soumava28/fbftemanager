import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { EditftedilougeComponent } from './editftedilouge.component';

describe('EditftedilougeComponent', () => {
  let component: EditftedilougeComponent;
  let fixture: ComponentFixture<EditftedilougeComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ EditftedilougeComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(EditftedilougeComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
